# -*- coding: utf-8 -*-
"""
Created on Mon Jan 19 09:54:42 2026

@author: thiago.onohara
"""

from dataclasses import dataclass
from datetime import date, datetime
from html import unescape
from html.parser import HTMLParser
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

try:
    from urllib3.util.retry import Retry
except Exception:  # pragma: no cover
    Retry = None


# =============================================================================
# Constants (from your investpy-like snippet)
# =============================================================================
from investing_constants import (
    COUNTRY_ID_FILTERS,
    CATEGORY_FILTERS,
    IMPORTANCE_RATINGS,
    TIME_FILTERS,
)


# =============================================================================
# Helpers
# =============================================================================

def _norm(s: str) -> str:
    return " ".join(s.strip().lower().split())


def _as_csv(values: Optional[Iterable[Union[str, int]]]) -> Optional[str]:
    if not values:
        return None
    return ",".join(str(v) for v in values)


def _tz_to_offset(time_zone: Optional[str]) -> str:
    """
    Aceita:
      - None -> "-03:00" (padrão BR; ajuste se quiser)
      - "-03:00" / "+00:00"
      - "GMT -3:00" (estilo investpy)
    Retorna sempre "±HH:MM".
    """
    if time_zone is None:
        return "-03:00"

    tz = time_zone.strip()
    if tz.startswith(("+", "-")) and len(tz) == 6 and tz[3] == ":":
        return tz

    if tz.upper().startswith("GMT"):
        rest = tz[3:].strip()  # "-3:00" / "+1:00" etc.
        if not rest:
            return "+00:00"
        sign = "+" if rest[0] == "+" else "-"
        hh, mm = rest[1:].strip().split(":")
        return f"{sign}{int(hh):02d}:{int(mm):02d}"

    return tz


def _iso_range_from_ddmmyyyy(
    from_date: Optional[str],
    to_date: Optional[str],
    tz_offset: str,
) -> Tuple[str, str]:
    """
    Converte dd/mm/yyyy em intervalo ISO com offset:
      start_date = YYYY-mm-ddT00:00:00.000±HH:MM
      end_date   = YYYY-mm-ddT23:59:59.999±HH:MM
    Se None/None: usa 'hoje'.
    """
    off = _tz_to_offset(tz_offset)

    def parse_d(s: str) -> date:
        return datetime.strptime(s, "%d/%m/%Y").date()

    if from_date is None and to_date is None:
        d0 = date.today()
        d1 = d0
    else:
        if from_date is None or to_date is None:
            raise ValueError("Se passar 'from_date', passe também 'to_date' (dd/mm/yyyy).")
        
        if isinstance(from_date, datetime):
            from_date = datetime.strftime(from_date, '%d/%m/%Y')
        if isinstance(to_date, datetime):
            to_date = datetime.strftime(to_date, '%d/%m/%Y')
            
        d0 = parse_d(from_date)
        d1 = parse_d(to_date)
        if d0 > d1:
            raise ValueError("to_date deve ser >= from_date.")

    start = f"{d0.isoformat()}T00:00:00.000{off}"
    end = f"{d1.isoformat()}T23:59:59.999{off}"
    return start, end


def _to_yyyy_mm_dd(value: Optional[Union[str, date, datetime]]) -> Optional[str]:
    """
    Normaliza datas para o formato YYYY-mm-dd.
    Aceita datetime/date, dd/mm/yyyy e yyyy-mm-dd.
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None

        for dt_fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(raw, dt_fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

    raise ValueError("Data invalida. Use datetime/date, 'dd/mm/yyyy' ou 'yyyy-mm-dd'.")


def _date_range_to_yyyy_mm_dd(
    from_date: Optional[Union[str, date, datetime]],
    to_date: Optional[Union[str, date, datetime]],
) -> Tuple[str, str]:
    """
    Converte intervalo de datas para formato YYYY-mm-dd.
    Se ambos forem None, usa hoje.
    """
    if from_date is None and to_date is None:
        day = date.today().strftime("%Y-%m-%d")
        return day, day

    if from_date is None or to_date is None:
        raise ValueError("Se passar 'from_date', passe tambem 'to_date'.")

    start = _to_yyyy_mm_dd(from_date)
    end = _to_yyyy_mm_dd(to_date)

    if start is None or end is None:
        raise ValueError("from_date/to_date invalidos.")

    if start > end:
        raise ValueError("to_date deve ser >= from_date.")

    return start, end


def _map_countries_to_ids(countries: Optional[Sequence[str]]) -> Optional[List[int]]:
    if not countries:
        return None
    ids: List[int] = []
    for c in countries:
        key = _norm(c)
        v = COUNTRY_ID_FILTERS.get(key)
        if v is not None:
            ids.append(int(v))
    return ids or None


def _coerce_country_ids(countries: Optional[Sequence[Union[str, int]]]) -> Optional[List[int]]:
    """
    Aceita paises por nome (lower case investpy-style), string numerica ou int.
    """
    if not countries:
        return None

    ids: List[int] = []
    for country in countries:
        if isinstance(country, int):
            ids.append(int(country))
            continue

        if isinstance(country, str):
            raw = country.strip()
            if not raw:
                continue

            if raw.isdigit():
                ids.append(int(raw))
                continue

            mapped_id = COUNTRY_ID_FILTERS.get(_norm(raw))
            if mapped_id is not None:
                ids.append(int(mapped_id))

    return ids or None


def _map_categories_to_api_slugs(categories: Optional[Sequence[str]]) -> Optional[List[str]]:
    """
    Mantive seu comportamento atual (passa slug direto),
    mas deixei o CATEGORY_FILTERS disponível caso você queira voltar a usar.
    """
    if not categories:
        return None

    out: List[str] = []
    for cat in categories:
        key = _norm(cat)

        # Você tinha forçado "mapped = key" (mantive)
        mapped = key  # ou: CATEGORY_FILTERS.get(key, key)
        slug = str(mapped).lstrip("_")
        out.append(slug)

    return out or None


def _build_requests_session(
    *,
    total_retries: int = 5,
    backoff_factor: float = 0.4,
    status_forcelist: Sequence[int] = (429, 500, 502, 503, 504),
) -> requests.Session:
    s = requests.Session()
    if Retry is None:
        return s

    retry = Retry(
        total=total_retries,
        connect=total_retries,
        read=total_retries,
        status=total_retries,
        backoff_factor=backoff_factor,
        status_forcelist=set(status_forcelist),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=50)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


class _HolidayHtmlTableParser(HTMLParser):
    """Parser simples para extrair linhas da tabela de feriados do HTML retornado."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: List[List[str]] = []
        self._current_row: List[str] = []
        self._current_cell: List[str] = []
        self._inside_cell = False

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        if tag == "tr":
            self._current_row = []
            return

        if tag in {"td", "th"}:
            self._inside_cell = True
            self._current_cell = []

    def handle_data(self, data: str) -> None:
        if self._inside_cell:
            self._current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"td", "th"} and self._inside_cell:
            raw_text = "".join(self._current_cell).replace("\xa0", " ")
            text = unescape(" ".join(raw_text.split())).strip()
            self._current_row.append(text)
            self._inside_cell = False
            self._current_cell = []
            return

        if tag == "tr" and self._current_row:
            self.rows.append(self._current_row)


def _parse_holiday_html_table(raw_html: str) -> pd.DataFrame:
    """
    Parseia o HTML da resposta do endpoint de feriados em colunas:
    date_raw, zone, exchange_name, holiday
    """
    out_cols = ["date_raw", "zone", "exchange_name", "holiday"]

    if not isinstance(raw_html, str) or not raw_html.strip():
        return pd.DataFrame(columns=out_cols)

    parser = _HolidayHtmlTableParser()
    parser.feed(raw_html)

    rows: List[List[str]] = []
    for row in parser.rows:
        if len(row) < 4:
            continue

        row4 = [str(x).strip() for x in row[:4]]

        if _norm(row4[0]) == "date" and _norm(row4[1]) == "country":
            continue

        rows.append(row4)

    if not rows:
        return pd.DataFrame(columns=out_cols)

    df = pd.DataFrame(rows, columns=out_cols)
    df["date_raw"] = df["date_raw"].replace("", pd.NA).ffill()
    return df


# =============================================================================
# Client
# =============================================================================

@dataclass(frozen=True)
class InvestingApiConfig:
    domain_id: int = 1
    timeout_s: int = 30
    verify_tls: bool = False  # mantém o default que você já vinha usando no wrapper
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )


class DataFrameParser:
    DEFAULT_TZ: str = "America/Sao_Paulo"

    @staticmethod
    def _parse_occurrence_time_sp_naive(
        s: pd.Series,
        tz: str = "America/Sao_Paulo",
    ) -> pd.Series:
        print("--- Parsing Occurrence Time -> Datetime ---")
        # utc=True cria tz-aware em UTC; depois tz_convert e remove tz (naive) :contentReference[oaicite:1]{index=1}
        return (
            pd.to_datetime(s, utc=True, errors="coerce")
            .dt.tz_convert(tz)
            .dt.tz_localize(None)
        )

    @staticmethod
    def _parse_country_id(s: pd.Series) -> pd.Series:
        print("--- Parsing Country ID -> Name ---")
        reversed_map = dict(zip(COUNTRY_ID_FILTERS.values(), COUNTRY_ID_FILTERS.keys()))
        return s.map(reversed_map)

    @staticmethod
    def _parse_event_name(df_: pd.DataFrame) -> pd.Series:
        print("--- Parsing Event Name ---")
        
        if 'event_cycle_suffix' in df_.columns:
            s = (
                df_["short_name"]
                + " "
                + df_["event_cycle_suffix"].fillna("")
                + " "
                + df_["reference_period"].apply(lambda r: f"({r})" if isinstance(r, str) else "")
            ).str.strip()
            return s
        
        else:
            print('"event_cycle_suffix" not in df_.columns')
            print('dataframe columns: ', df_.columns)
            s = (
                df_["short_name"]
                + " "
                + df_["reference_period"].apply(lambda r: f"({r})" if isinstance(r, str) else "")
            ).str.strip()
            return s
            
class InvestingAPIClient:
    BASE_URL = "https://endpoints.investing.com/pd-instruments/v1"
    HOLIDAY_URL = "https://www.investing.com/holiday-calendar/Service/getCalendarFilteredData"

    def __init__(self, config: Optional[InvestingApiConfig] = None, session: Optional[requests.Session] = None):
        self.config = config or InvestingApiConfig()
        self.session = session or _build_requests_session()

    def _headers(self, extra: Optional[Mapping[str, str]] = None) -> Dict[str, str]:
        h = {
            "Accept": "*/*",
            "Origin": "https://www.investing.com",
            "Referer": "https://www.investing.com/",
            "User-Agent": self.config.user_agent,
        }
        if extra:
            h.update(dict(extra))
        return h

    def get_economic_occurrences(
        self,
        *,
        start_date_iso: str,
        end_date_iso: str,
        country_ids: Optional[Sequence[int]] = None,
        categories: Optional[Sequence[str]] = None,
        limit: int = 200,
        extra_params: Optional[Mapping[str, Any]] = None,
        headers: Optional[Mapping[str, str]] = None,
    ) -> Dict[str, Any]:
        url = f"{self.BASE_URL}/calendars/economic/events/occurrences"

        params: Dict[str, Any] = {
            "domain_id": str(self.config.domain_id),
            "limit": str(int(limit)),
            "start_date": start_date_iso,
            "end_date": end_date_iso,
        }
        if country_ids:
            params["country_ids"] = _as_csv(country_ids)
        if categories:
            params["categories"] = _as_csv(categories)
        if extra_params:
            params.update(dict(extra_params))

        resp = self.session.get(
            url,
            headers=self._headers(headers),
            params=params,
            timeout=self.config.timeout_s,
            verify=self.config.verify_tls,
        )
        resp.raise_for_status()
        return resp.json()

    # =============================================================================
    # NEW USAGE: investing = InvestingAPIClient(); investing.economic_calendar(...)
    # =============================================================================
    def get_economic_calendar(
        self,
        time_zone: Optional[str] = "GMT -3:00",
        time_filter: str = "time_only",  # compat
        countries: Optional[List[str]] = None,
        importances: Optional[List[str]] = None,  # (ainda não usado; mantido p/ compat)
        categories: Optional[List[str]] = None,
        from_date: Optional[str] = None,  # dd/mm/yyyy
        to_date: Optional[str] = None,    # dd/mm/yyyy
        *,
        limit: int = 300,
        extra_params: Optional[Mapping[str, Any]] = None,
    ) -> pd.DataFrame:
        if time_filter not in TIME_FILTERS:
            raise ValueError(f"time_filter inválido: {time_filter}. Use {list(TIME_FILTERS.keys())}")

        start_iso, end_iso = _iso_range_from_ddmmyyyy(
            from_date, to_date, tz_offset=_tz_to_offset(time_zone)
        )

        country_ids = _map_countries_to_ids(countries)
        category_slugs = _map_categories_to_api_slugs(categories)

        raw = self.get_economic_occurrences(
            start_date_iso=start_iso,
            end_date_iso=end_iso,
            country_ids=country_ids,
            categories=category_slugs,
            limit=limit,
            extra_params=extra_params,
        )

        print("TOTAL EVENTS (dict keys): ", len(raw), sep='\n'*3)
        print('EVENTS: ', pd.json_normalize(raw['events']), sep='\n'*3)
        print('OCCURENCES: ', pd.json_normalize(raw['occurrences']), sep='\n'*3)

        raw_event_df = pd.json_normalize(raw["events"])
        raw_occ_df = pd.json_normalize(raw["occurrences"])

        # Merge Events and Occurrences
        df = raw_event_df.merge(raw_occ_df, on=["event_id"])

        # Parsing
        df["id"] = df["event_id"]
        df["datetime"] = DataFrameParser._parse_occurrence_time_sp_naive(df["occurrence_time"])
        df["date"] = df["datetime"].dt.strftime("%d/%m/%Y")
        df["time"] = df["datetime"].dt.strftime("%H:%M")
        df["zone"] = DataFrameParser._parse_country_id(df["country_id"])
        df["event"] = DataFrameParser._parse_event_name(df)
        if 'actual' in df.columns:
            df['actual'] = df['actual'].fillna('').astype(str) + df['unit'].fillna('').astype(str)
        base_cols = pd.Index(["id", "date", "time", "zone", "currency", "importance", "event", "actual", "forecast", "previous"])
        matched_cols = df.columns.intersection(base_cols)
        unmatched_cols = df.columns.difference(base_cols)

        print('MATCHED COLUMNS: ', matched_cols)
        print('UNMATCHED COLUMNS: ', unmatched_cols)
        return df[matched_cols]

    def get_holiday_calendar(
        self,
        countries: Optional[List[Union[str, int]]] = None,
        from_date: Optional[Union[str, date, datetime]] = None,
        to_date: Optional[Union[str, date, datetime]] = None,
        *,
        limit_from: int = 0,
        extra_payload: Optional[Mapping[str, Any]] = None,
        headers: Optional[Mapping[str, str]] = None,
    ) -> pd.DataFrame:
        """
        Busca feriados de mercado no endpoint legado da pagina Holiday Calendar.

        Endpoint:
        https://www.investing.com/holiday-calendar/Service/getCalendarFilteredData
        """
        start_date, end_date = _date_range_to_yyyy_mm_dd(from_date, to_date)
        country_ids = _coerce_country_ids(countries)

        holiday_headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://www.investing.com/holiday-calendar/",
        }
        if headers:
            holiday_headers.update(dict(headers))
        base_cols = ["id", "date", "time", "zone", "currency", "importance", "event", "actual", "forecast", "previous"]

        def fetch_single_country(country_id: Optional[int]) -> pd.DataFrame:
            payload: Dict[str, Any] = {
                "dateFrom": start_date,
                "dateTo": end_date,
                "country": "" if country_id is None else str(int(country_id)),
                "currentTab": "custom",
                "submitFilters": "1",
                "limit_from": str(int(limit_from)),
            }

            if extra_payload:
                payload.update(dict(extra_payload))

            resp = self.session.post(
                self.HOLIDAY_URL,
                headers=self._headers(holiday_headers),
                data=payload,
                timeout=self.config.timeout_s,
                verify=self.config.verify_tls,
            )
            resp.raise_for_status()

            raw_html = ""
            try:
                parsed = resp.json()
                if isinstance(parsed, Mapping):
                    raw_html = str(parsed.get("data", "") or "")
            except ValueError:
                raw_html = resp.text

            parsed_df = _parse_holiday_html_table(raw_html)

            if parsed_df.empty:
                return pd.DataFrame(columns=base_cols)

            parsed_df["date_raw"] = parsed_df["date_raw"].replace("", pd.NA).ffill()
            parsed_df["datetime"] = pd.to_datetime(parsed_df["date_raw"], errors="coerce")
            parsed_df = parsed_df.dropna(subset=["datetime"])

            if parsed_df.empty:
                return pd.DataFrame(columns=base_cols)

            requested_start = pd.to_datetime(start_date)
            requested_end = pd.to_datetime(end_date)
            parsed_df = parsed_df[parsed_df["datetime"].between(requested_start, requested_end)]

            if parsed_df.empty:
                return pd.DataFrame(columns=base_cols)

            parsed_df["zone"] = (
                parsed_df["zone"]
                .fillna("")
                .astype(str)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )
            parsed_df["event"] = parsed_df["holiday"].fillna("").astype(str).str.strip()
            parsed_df.loc[parsed_df["event"].eq(""), "event"] = "Market Holiday"

            parsed_df["date"] = parsed_df["datetime"].dt.strftime("%d/%m/%Y")
            parsed_df["time"] = "00:00"
            parsed_df["currency"] = ""
            parsed_df["importance"] = "low"
            parsed_df["actual"] = ""
            parsed_df["forecast"] = ""
            parsed_df["previous"] = ""

            parsed_df["id"] = (
                "holiday_"
                + parsed_df["datetime"].dt.strftime("%Y%m%d")
                + "_"
                + parsed_df["zone"].str.lower().str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_")
                + "_"
                + parsed_df["event"].str.lower().str.replace(r"[^a-z0-9]+", "_", regex=True).str.strip("_")
            )

            return parsed_df[base_cols].drop_duplicates().reset_index(drop=True)

        country_ids_to_fetch = list(dict.fromkeys(country_ids)) if country_ids else [None]
        frames = [fetch_single_country(country_id) for country_id in country_ids_to_fetch]
        frames = [frame for frame in frames if not frame.empty]

        if not frames:
            return pd.DataFrame(columns=base_cols)

        return pd.concat(frames, ignore_index=True).drop_duplicates().reset_index(drop=True)


# =============================================================================
# Example
# =============================================================================
if __name__ == "__main__":
    investing = InvestingAPIClient()  # <- novo modo de uso

    df = investing.get_economic_calendar(
        countries=[
            "brazil", "united states", "euro zone", "canada", "united kingdom", "canada",
            "china", "sweden", "norway", "south korea", "mexico", "chile", "colombia",
            "south africa", "czech republic", "hungary", "poland", "singapore", "japan",
            "switzerland", "france", "spain", "italy"
        ],
        categories=["inflation", "employment", "central_banks"],
        from_date="19/01/2026",
        to_date="20/01/2026",
    )

    print(df.head(30))
