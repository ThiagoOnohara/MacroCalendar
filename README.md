# Economic Calendar → Outlook (local + Outlook)

Projeto para **buscar eventos econômicos** (Investing via `investpy`), **filtrar/deduplicar** e **publicar automaticamente** no seu **calendário do Outlook**. Pensado para rodar **localmente no Windows** e integrar com **Excel** (via `xlwings`).

> **Importante**: este projeto **não** é um serviço na nuvem. Ele depende do Outlook instalado na sua máquina (COM Automation) e, opcionalmente, do Excel para visualizar/ajustar os eventos antes de enviar ao calendário.

---

## Sumário
- [Arquitetura](#arquitetura)
- [Recursos](#recursos)
- [Pré‑requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como usar](#como-usar)
  - [1) Pela planilha (Excel + xlwings)](#1-pela-planilha-excel--xlwings)
  - [2) Pela linha de comando (Python)](#2-pela-linha-de-comando-python)
  - [3) Agendamento (Windows Task Scheduler)](#3-agendamento-windows-task-scheduler)
- [Personalizações úteis](#personalizações-úteis)
- [Estrutura de pastas](#estrutura-de-pastas)
- [Solução de problemas (FAQ)](#solução-de-problemas-faq)
- [Avisos legais](#avisos-legais)

---

## Arquitetura

- **`eco_calendar.py`**: wrapper simples em torno do `investpy` para coletar o calendário econômico por **intervalo de datas**, **categorias** e **países**.
- **`calendar_events.py`**: carrega, filtra e deduplica os eventos; permite **exportar para Excel** e **criar/atualizar compromissos** no Outlook (via `pywin32`). Contém utilitários (merge, tokenização/semelhança rudimentar) e os endpoints de execução (`export_events`, `add_to_agenda`, `_run_scheduled_task`).

Fluxo típico:
1. Buscar calendário econômico (inflation, employment, central_banks, activity).
2. Pré‑processar (normalização de nomes, remoção de duplicatas, filtros por relevância/termos).
3. (Opcional) Enviar a **aba `EVENTS`** no Excel para revisão.
4. Publicar no **Calendário do Outlook** (subpasta "FX" por padrão), com **categorias de cor** por tipo de evento e **lembrete**.

---

## Recursos
- Coleta de eventos por **país** e **categoria** para os próximos *N* dias.
- **Filtros manuais** por termos (mantém relevantes e exclui ruído).
- **Deduplicação leve** por similaridade de tokens no mesmo horário/zona.
- Exportação para **Excel** e importação de volta para criar compromissos.
- Integração **Outlook (COM)**: cria/atualiza compromissos, evita duplicados, define lembrete e categorias.

---

## Pré‑requisitos
- **Sistema**: Windows 10/11 com **Outlook** instalado e perfil configurado.
- **Python**: 3.9–3.12.
- **Excel** (opcional, para revisão via planilha).

### Dependências (pip)
Crie um `requirements.txt` como abaixo:
```txt
pandas
xlwings
investpy
pywin32
```
> Bibliotecas da *standard library* (ex.: `os`, `re`, `datetime`, `functools`, `collections`) **não** entram no requirements.

Instale com:
```bash
pip install -r requirements.txt
```

---

## Instalação
1. Clone este repositório.
2. Crie/ative um ambiente virtual (recomendado).
3. Instale as dependências via `requirements.txt`.
4. Garanta que o **Outlook** abre normalmente com sua conta.

---

## Configuração
Abra `calendar_events.py` e ajuste os pontos abaixo, se necessário:

1) **Subpasta do calendário no Outlook**
```python
my_calendar_name = 'FX'  # nome da subpasta do seu calendário
```
Crie essa subpasta no Outlook (Calendário → Clique direito → Novo Calendário) ou troque o nome para uma pasta que já exista.

2) **Domínio de e-mail para Required Attendees**
```python
mail_adress = '@youradress.com'
# usa os logins locais + domínio para formar os e-mails
```
Se não deseja adicionar participantes, deixe a lista vazia.

3) **Categorias de cor** (Outlook)
Mapeadas em `category_to_color` dentro de `add_to_agenda()`:
```python
{
  'cb': 'Dark',
  'inflation': 'Green',
  'activity': 'Blue',
  'employ': 'Red'
}
```
Certifique‑se de que essas categorias existem no Outlook (ou ajuste para categorias existentes na sua instalação).

4) **Regiões/países e códigos**
Os dicionários `country_total`, `country_cb`, `country_employment` e o `zone_to_code` podem ser ajustados conforme seu universo de países.

5) **User‑Agent do investpy**
Em `eco_calendar.py` define‑se um `user_agent` customizado (header HTTP). Mantenha atualizado caso o `investpy`/Investing.com mude políticas de acesso.

---

## Como usar

### 1) Pela planilha (Excel + xlwings)
- Prepare um arquivo Excel com uma **aba chamada `EVENTS`**.
- Com o Excel aberto, chame a macro do `xlwings` (ou um botão associado) que dispara `export_events()`:
  - Ele preenche a partir da célula **A5** uma tabela com colunas como `datetime`, `zone`, `event`, `category`, `importance`, etc.
- Revise/edite, se quiser, e então execute `add_to_agenda()` para criar/atualizar os compromissos no Outlook.

> Dica: No `calendar_events.py`, `export_events()` e `add_to_agenda()` já estão decorados com `@xw.sub`, então podem ser conectados diretamente a botões no Excel via xlwings.

### 2) Pela linha de comando (Python)
Execute diretamente para rodar o fluxo de tarefa:
```bash
python calendar_events.py
```
Isso chama `_run_scheduled_task()`, que:
1. Executa `add_to_agenda()` (coleta e envia os eventos para o Outlook).
2. Tenta enviar um e‑mail de log (`send_logging_email()`), ajuste o destinatário em `mail.To` se quiser realmente usar.

### 3) Agendamento (Windows Task Scheduler)
- **Ação**: `python`
- **Argumentos**: caminho completo para `calendar_events.py`
- **Iniciar em**: pasta do projeto
- **Agende** conforme sua necessidade (ex.: diariamente às 06:00). O script usará o Outlook do usuário logado.

---

## Personalizações úteis
- **Janela de busca de eventos**: em `get_and_filter_events(next_days=10)` altere `next_days`.
- **Deduplicação**: ajuste `threshold`/`top_n_words` em `dedupe_top`.
- **Filtros manuais**: refine as expressões em `get_calendar_from_investing()` para suas preferências (incluir/excluir termos por categoria).
- **Lembrete e duração**: variáveis `remind_before` (min) e duração fixa de 30min (em `add_to_agenda`).
- **Pasta do calendário**: troque `object_type='my_calendar'` para `'calendar'` se quiser usar a pasta padrão em vez de uma subpasta.

---

## Estrutura de pastas
```
.
├─ calendar_events.py     # integra Excel/Outlook, filtros e publicação
├─ eco_calendar.py        # coleta de eventos via investpy
├─ requirements.txt       # pandas, xlwings, investpy, pywin32
└─ README.md              # este arquivo
```

---

## Solução de problemas (FAQ)
**1) `investpy` retornou DataFrame vazio**  
- Verifique o intervalo de datas, categorias e países.
- O Investing.com pode bloquear temporariamente acessos automatizados. Tente novamente e mantenha o `user_agent` atualizado.

**2) Erro COM/Outlook (ex.: não abre, permissões)**  
- Confirme que o Outlook abre normalmente com sua conta.
- Execute o Python com o mesmo usuário do Outlook.

**3) Compromissos duplicados**  
- A função `find_existing` tenta evitar duplicados (mesmo assunto e mesma data). Se mudar o *subject* (ex.: renomear a categoria ou `zone_to_code`), pode não reconhecer o item antigo. Ajuste a lógica se necessário.

**4) Categorias não aparecem no Outlook**  
- As categorias usadas precisam existir no Outlook do usuário. Crie/rename conforme sua instalação.

**5) Fuso horário**  
- Os `datetime` enviados são *naive* (string `YYYY-MM-DD HH:MM`). Garanta que seus horários já estejam no fuso correto antes do envio.

**6) Integração com Excel**  
- A planilha precisa ter a aba `EVENTS` e o intervalo a partir de `A5`.
- `xw.Book.caller()` só funciona quando o Python é chamado via xlwings/Excel; fora disso o código entra no *fallback* e usa o DataFrame diretamente.

---

## Avisos legais
- `investpy` usa dados do Investing.com e pode estar sujeito a mudanças de layout/limites de acesso. Use com responsabilidade.
- O projeto é fornecido **sem garantias**. Revise os eventos no seu calendário antes de compartilhar.

---

**Pronto!** Ajuste as configurações acima, rode localmente e mantenha seu calendário macroeconômico sempre em dia no Outlook. ☕📅