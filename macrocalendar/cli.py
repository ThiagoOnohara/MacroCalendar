"""Command-line interface for MacroCalendar."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import replace
from pathlib import Path
from typing import Optional, Sequence

import requests

from .config import AppConfig, default_config_path
from .outlook import OutlookClient, OutlookError


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="macrocalendar",
        description="Busca eventos macroeconômicos e sincroniza com o Outlook Classic.",
    )
    parser.add_argument("--version", action="version", version="%(prog)s 2.0.0")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser(
        "doctor",
        help="verifica Python, dependências, Outlook Classic e calendários",
    )

    config_parser = subparsers.add_parser("config", help="cria ou mostra a configuração")
    config_parser.add_argument("action", choices=("init", "show"), nargs="?", default="show")
    config_parser.add_argument("--path", type=Path, help=argparse.SUPPRESS)

    for command, help_text in (
        ("preview", "busca e exibe os próximos eventos"),
        ("sync", "busca eventos e publica no Outlook Classic"),
    ):
        command_parser = subparsers.add_parser(command, help=help_text)
        command_parser.add_argument("--days", type=int, help="quantos dias consultar")
        command_parser.add_argument("--calendar-name", help="nome do calendário no Outlook")
        command_parser.add_argument(
            "--no-holidays",
            action="store_true",
            help="não consultar o endpoint legado de feriados",
        )
        if command == "preview":
            command_parser.add_argument("--csv", type=Path, help="salva o resultado em CSV")
        else:
            command_parser.add_argument(
                "--dry-run",
                action="store_true",
                help="mostra o que seria criado/atualizado sem salvar",
            )
            command_parser.add_argument(
                "--non-interactive",
                action="store_true",
                help="falha se --calendar-name não estiver configurado",
            )

    return parser


def _load_config() -> AppConfig:
    config = AppConfig.load()
    config.validate()
    return config


def _apply_overrides(config: AppConfig, args) -> AppConfig:
    values = {}
    if getattr(args, "days", None) is not None:
        values["days"] = args.days
    if getattr(args, "calendar_name", None):
        values["calendar_name"] = args.calendar_name
    if getattr(args, "no_holidays", False):
        values["include_holidays"] = False
    result = replace(config, **values)
    result.validate()
    return result


def _choose_calendar(config: AppConfig, args) -> OutlookClient:
    client = OutlookClient(config.calendar_name).connect()
    if config.calendar_name or getattr(args, "non_interactive", False):
        return client

    names = client.calendar_names()
    print("\nCalendários encontrados:")
    for index, name in enumerate(names, start=1):
        print(f"  {index}. {name}")
    answer = input("Escolha um calendário (Enter usa o padrão): ").strip()
    if not answer:
        return client
    try:
        selected = names[int(answer) - 1]
    except (ValueError, IndexError) as exc:
        raise OutlookError("Escolha inválida. Execute novamente e informe um número da lista.") from exc

    return OutlookClient(selected).connect()


def _fetch(config: AppConfig):
    from calendar_events import get_and_filter_events

    print(f"Consultando Investing.com ({config.days} dias)...")
    try:
        events = get_and_filter_events(
            config.days,
            include_holidays=config.include_holidays,
            countries=config.countries,
        categories=config.categories,
        timezone=config.timezone,
        verify_tls=config.verify_tls,
    )
    except requests.exceptions.SSLError as exc:
        raise RuntimeError(
            "O Investing.com apresentou um erro de certificado SSL. "
            "Se sua rede exigir ignorar essa validação, altere "
            "'verify_tls' para false em %APPDATA%\\MacroCalendar\\config.json."
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Falha ao consultar/processar o Investing.com: {exc}") from exc
    print(f"Eventos selecionados: {len(events)}")
    return events


def _preview(args) -> int:
    config = _apply_overrides(_load_config(), args)
    events = _fetch(config)
    if getattr(args, "csv", None):
        args.csv.parent.mkdir(parents=True, exist_ok=True)
        events.to_csv(args.csv, index=False, encoding="utf-8-sig")
        print(f"CSV salvo em: {args.csv}")
    elif events.empty:
        print("Nenhum evento encontrado.")
    else:
        columns = [column for column in ("datetime", "zone", "event", "category", "importance") if column in events]
        print(events[columns].to_string(index=False))
    return 0


def _sync(args) -> int:
    config = _apply_overrides(_load_config(), args)
    if not config.calendar_name and args.non_interactive:
        raise OutlookError(
            "Informe --calendar-name ou configure calendar_name antes de usar --non-interactive."
        )

    # Validate Outlook before making the potentially slow data requests.
    print("Verificando Outlook Classic...")
    client = _choose_calendar(config, args)
    print(f"Calendário selecionado: {client.calendar_name}")
    events = _fetch(config)

    if events.empty:
        print("Nenhum evento para sincronizar.")
        return 0

    from calendar_events import sync_events

    mode = " (dry-run)" if args.dry_run else ""
    print(f"Sincronizando {len(events)} eventos{mode}...")
    counts = sync_events(events, config=config, dry_run=args.dry_run, client=client)
    print(
        "Resumo: "
        f"{counts['created']} criados, "
        f"{counts['updated']} atualizados, "
        f"{counts['duplicates']} com múltiplas correspondências."
    )
    return 0


def _doctor() -> int:
    print("MacroCalendar doctor")
    print(f"Python: {sys.version.split()[0]} ({sys.executable})")
    if sys.platform != "win32":
        print("ERRO: a publicação no Outlook requer Windows.")
        return 1

    try:
        import pandas  # noqa: F401
        import requests  # noqa: F401
    except ImportError as exc:
        print(f"ERRO: dependência Python ausente: {exc.name}")
        return 1
    print("Dependências HTTP e dados: OK")

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        print("ERRO: pywin32 não está instalado.")
        return 1
    print("pywin32: OK")

    try:
        config = AppConfig.load()
        client = OutlookClient(config.calendar_name).connect()
        print(f"Outlook Classic/COM: OK ({client.calendar_name})")
        print("Calendários:")
        for name in client.calendar_names():
            print(f"  - {name}")
    except (OutlookError, ValueError) as exc:
        print(f"ERRO: {exc}")
        return 1
    return 0


def _config(args) -> int:
    path = args.path or default_config_path()
    if args.action == "init":
        if path.exists():
            answer = input(f"A configuração já existe em {path}. Sobrescrever? [y/N] ").strip().lower()
            if answer not in {"y", "yes", "s", "sim"}:
                print("Operação cancelada.")
                return 0
        saved = AppConfig().save(path)
        print(f"Configuração criada em: {saved}")
        print("Edite calendar_name se quiser evitar a seleção interativa.")
    else:
        config = AppConfig.load(path)
        print(f"Arquivo: {path}")
        print(json.dumps(config.__dict__, ensure_ascii=False, indent=2))
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            return _doctor()
        if args.command == "config":
            return _config(args)
        if args.command == "preview":
            return _preview(args)
        if args.command == "sync":
            return _sync(args)
    except KeyboardInterrupt:
        print("\nOperação cancelada.", file=sys.stderr)
        return 130
    except (OutlookError, ValueError, OSError, KeyError, RuntimeError) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"ERRO inesperado ({type(exc).__name__}): {exc}", file=sys.stderr)
        return 1
    return 1
