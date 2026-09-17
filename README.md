# MacroCalendar

CLI para buscar eventos macroeconômicos do Investing.com e criar ou atualizar compromissos no Outlook Classic para Windows.

```text
Investing.com → filtros → MacroCalendar → Outlook Classic
```

O projeto consulta o endpoint HTTP usado pelo calendário do Investing.com. Ele não usa mais `investpy` nem Excel/xlwings.

## Requisitos

- Windows 10/11;
- Python 3.10 ou superior;
- Outlook Classic instalado, aberto e conectado a um perfil;
- acesso à internet.

O New Outlook não oferece suporte à automação COM usada para escrever no calendário. Se o New Outlook estiver aberto, abra também o Outlook Classic ou volte para ele.

## Instalação

Na pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

Também é possível executar sem instalar o entry point:

```powershell
python -m macrocalendar --help
```

## Primeiro uso

Execute o diagnóstico:

```powershell
macrocalendar doctor
```

Ele verifica o Python, dependências, acesso ao Outlook Classic e lista os calendários encontrados.

Para criar uma configuração inicial:

```powershell
macrocalendar config init
```

O arquivo fica em `%APPDATA%\MacroCalendar\config.json`. Ele não contém senha nem token. Para escolher o calendário diretamente, edite:

```json
{
  "calendar_name": "Macro",
  "days": 10,
  "timezone": "America/Sao_Paulo",
  "verify_tls": false,
  "reminder_minutes": 15,
  "duration_minutes": 30,
  "include_holidays": true
}
```

Se `calendar_name` ficar vazio, o CLI mostrará os calendários e perguntará qual usar.

## Uso

Ver os eventos sem alterar o Outlook:

```powershell
macrocalendar preview
macrocalendar preview --days 5
macrocalendar preview --csv eventos.csv
```

Sincronizar:

```powershell
macrocalendar sync
macrocalendar sync --calendar-name Macro
macrocalendar sync --days 14 --dry-run
macrocalendar sync --no-holidays
```

O `sync` valida o Outlook antes de consultar os dados e informa claramente em qual etapa ocorreu um erro. Eventos existentes são atualizados quando o assunto e o horário coincidem; eventos duplicados não são apagados automaticamente.

Para automação sem perguntas interativas:

```powershell
macrocalendar sync --calendar-name Macro --non-interactive
```

## Task Scheduler

O Outlook COM depende de uma sessão interativa. Configure a tarefa para executar somente quando o usuário estiver conectado, usando o Python do ambiente virtual:

```text
Ação: C:\caminho\do\projeto\.venv\Scripts\python.exe
Argumentos: -m macrocalendar sync --calendar-name Macro --non-interactive
Iniciar em: C:\caminho\do\projeto
```

## Configuração padrão

O programa acompanha inflação, emprego, bancos centrais, atividade econômica e feriados para um conjunto de países relevantes. A lista completa pode ser ajustada em `config.json` usando `countries` e `categories`.

Exemplo:

```json
{
  "calendar_name": "Macro",
  "countries": ["brazil", "united states", "euro zone"],
  "categories": ["inflation", "employment", "central_banks", "economic_activity"]
}
```

## Estrutura

- `macrocalendar/cli.py`: comandos e mensagens do CLI;
- `macrocalendar/config.py`: configuração persistente;
- `macrocalendar/outlook.py`: integração segura com Outlook Classic;
- `investing_calendar.py`: cliente HTTP e normalização da resposta;
- `calendar_events.py`: filtros, deduplicação e compatibilidade com o uso antigo;
- `investing_constants.py`: mapas de países e constantes do endpoint.

`eco_calendar.py` permanece apenas como uma pequena camada de compatibilidade para scripts antigos. A implementação interna não depende mais de `investpy`.

## Limitações

- A integração com o Investing.com usa endpoints sujeitos a mudanças de layout, parâmetros e bloqueios;
- Outlook New não é suportado pela integração COM;
- os horários são normalizados para `America/Sao_Paulo` antes de serem gravados no Outlook;
- o programa cria compromissos, não reuniões, e não envia convites.
