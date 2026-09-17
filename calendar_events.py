# -*- coding: utf-8 -*-
"""
@author: thiago.onohara
"""

import pandas as pd
from functools import reduce
from datetime import datetime, timedelta
import re
from investing_calendar import InvestingAPIClient, InvestingApiConfig
from macrocalendar.config import AppConfig
from macrocalendar.outlook import OutlookClient

calendar = InvestingAPIClient()

#%% UTILS
def merge_dfs(dfs, on=[]):
  """
  Merges a list of DataFrames using reduce.

  Args:
    dfs: A list of DataFrames to merge.

  Returns:
    A single DataFrame containing all the data from the input DataFrames.
  """

  # Use reduce to iteratively merge DataFrames
  merged_df = reduce(lambda left, right: pd.merge(left=left, right=right, on=on, how='inner'), dfs)
  return merged_df

#%% Eco Data

country_total = {
    'france':'fra',
    'germany':'ger',
    'italy':'itl',
    'sweden':'sek',
    'spain':'spn',
    'czech republic':'czk',
    'hungary':'huf', 
    'poland':'pln',
    'united kingdom':'gbp',
    'brazil':'brl',
    'colombia':'cop',
    'chile':'clp',
    'mexico':'mxn',
    'india':'inr',
    'south korea':'krw', 
    'south africa':'zar',
    'china':'cnh',
    'united states':'usd',
    'new zealand':'nzd',
    'japan':'jpy',
    'australia':'aud',
    'canada':'cad',
    'norway':'nok',  
    'switzerland':'chf',
    'euro zone':'eur',
    'turkey':'try'
    }

country_cb = {
    'sweden':'sek',
    'hungary':'huf', 
    'poland':'pln',
    'united kingdom':'gbp',
    'brazil':'brl',
    'colombia':'cop',
    'chile':'clp',
    'mexico':'mxn',
    'south korea':'krw', 
    'south africa':'zar',
    'china':'cnh',
    'united states':'usd',
    'new zealand':'nzd',
    'japan':'jpy',
    'australia':'aud',
    'canada':'cad',
    'norway':'nok',  
    'switzerland':'chf',
    'euro zone':'eur'
    }

country_employment = {
    'sweden':'sek',
    'united kingdom':'gbp',
    'brazil':'brl',
    'mexico':'mxn',
    'china':'cnh',
    'united states':'usd',
    'new zealand':'nzd',
    'japan':'jpy',
    'australia':'aud',
    'canada':'cad',
    'norway':'nok',  
    'switzerland':'chf',
    'euro zone':'eur'
    }

from collections import Counter
def tokenize(text):
    return re.findall(r'\b\w+\b', text.lower())

def dedupe_top(df, threshold=0.7, top_n_words=15):
    df = df.copy().sort_values(by=['zone', 'date', 'time', 'event'])
    df['group_key'] = df['zone'].astype(str) + '_' + df['date'] + '_' + df['time']
    
    def filter_group(group):
        # Contar palavras mais frequentes no grupo
        if len(group['event_lower']) > 1:
            word_counter = Counter()
            for event in group['event_lower']:
                tokens = tokenize(event)
                word_counter.update(tokens)
            top_words = set([w for w, _ in word_counter.most_common(top_n_words)])
            
            # Selecionar evento base (primeiro)
            base_tokens = set(tokenize(group.iloc[0]['event_lower'])) & top_words
            keep_rows = [True]  # Mantemos o primeiro sempre
    
            for idx in range(1, len(group)):
                current_tokens = set(tokenize(group.iloc[idx]['event_lower'])) & top_words
                if len(base_tokens) == 0:
                    similarity = 0
                else:
                    similarity = len(base_tokens & current_tokens) / len(base_tokens)
                keep_rows.append(similarity < threshold)  # Só mantém se for diferente
    
            return group[keep_rows]
        else: 
            return group

    data_columns = [column for column in df.columns if column != 'group_key']
    result = df.groupby('group_key', group_keys=False)[data_columns].apply(filter_group)
    return result.reset_index(drop=True).sort_values(by=['date', 'time', 'zone'])

def _timezone_to_investing_offset(timezone):
    """Return the current IANA timezone offset in Investing's GMT format."""
    try:
        offset = pd.Timestamp.now(tz=timezone).utcoffset()
    except Exception as exc:
        raise ValueError(f"Timezone inválido: {timezone}") from exc
    total_minutes = int(offset.total_seconds() // 60)
    sign = '+' if total_minutes >= 0 else '-'
    total_minutes = abs(total_minutes)
    hours, minutes = divmod(total_minutes, 60)
    return f"GMT {sign}{hours}:{minutes:02d}"


def get_calendar_from_investing(
    next_days=7,
    include_holidays=True,
    countries=None,
    categories=None,
    timezone='America/Sao_Paulo',
    verify_tls=None,
):
    countries_inflation = list(country_total.keys())
    countries_central_banks = list(country_cb.keys())
    countries_employment = list(country_employment.keys())
    countries_activity = list(country_total.keys())
    countries_holidays = list(country_total.keys())

    if countries:
        countries_inflation = countries_central_banks = countries_employment = countries_activity = list(countries)
        countries_holidays = list(countries)

    requested_categories = {
        str(category).strip().lower() for category in (categories or [
            'inflation', 'employment', 'central_banks', 'economic_activity'
        ])
    }
    if 'activity' in requested_categories:
        requested_categories.add('economic_activity')

    investing_timezone = _timezone_to_investing_offset(timezone)
    api_calendar = calendar
    current_config = getattr(calendar, 'config', None)
    current_verify_tls = getattr(current_config, 'verify_tls', verify_tls)
    if verify_tls is not None and current_verify_tls != verify_tls:
        api_calendar = InvestingAPIClient(InvestingApiConfig(verify_tls=verify_tls))

    empty_events = lambda: pd.DataFrame(columns=[
        'id', 'date', 'time', 'zone', 'currency',
        'importance', 'event', 'actual', 'forecast', 'previous'
    ])

    def fetch(category, country_list):
        if category not in requested_categories:
            return empty_events()
        return api_calendar.get_economic_calendar(
            from_date=datetime.today(),
            to_date=datetime.today()+timedelta(days=next_days),
            time_zone=investing_timezone,
            categories=[category],
            countries=country_list,
            output_timezone=timezone)

    data_inflation = fetch('inflation', countries_inflation)
    data_employ = fetch('employment', countries_employment)
    data_cb = fetch('central_banks', countries_central_banks)
    data_activity = fetch('economic_activity', countries_activity)

    # Market holidays use a legacy Investing endpoint and can be disabled.
    if include_holidays:
        try:
            data_holiday = api_calendar.get_holiday_calendar(
                from_date=datetime.today(),
                to_date=datetime.today()+timedelta(days=next_days),
                countries=countries_holidays)
        except Exception as error:
            print('Holiday calendar request failed:', error)
            data_holiday = pd.DataFrame(columns=[
                'id', 'date', 'time', 'zone', 'currency',
                'importance', 'event', 'actual', 'forecast', 'previous'
                ])
    else:
        data_holiday = pd.DataFrame(columns=[
            'id', 'date', 'time', 'zone', 'currency',
            'importance', 'event', 'actual', 'forecast', 'previous'
            ])
    
    def pre_filter_events(df):
        base_drop_cols = ['id', 'actual', 'forecast', 'previous']
        matched_drop_cols = df.columns.intersection(pd.Index(base_drop_cols))
        data_not_duplt = df.copy().drop(matched_drop_cols, axis=1)
        data_not_duplt['event_lower'] = data_not_duplt['event'].str.split('(').str[0].str.strip().str.lower()
        data_not_duplt_ = data_not_duplt.drop_duplicates(subset=['date', 'time', 'zone', 'event_lower'], keep='first')
        return data_not_duplt_
    
    data_inflation_ = pre_filter_events(data_inflation)
    data_employ_ = pre_filter_events(data_employ)
    data_cb_ = pre_filter_events(data_cb)
    data_activity_ = pre_filter_events(data_activity)
    data_holiday_ = pre_filter_events(data_holiday)
    
    #MANUAL FILTERINGS
    data_inflation_manual = data_inflation_[
        (data_inflation_['importance'].isin(['medium', 'high']))|
        (data_inflation_['event_lower'].str.contains('cpi'))|
        (data_inflation_['event_lower'].str.contains('ppi'))|
        (data_inflation_['event_lower'].str.contains('pce'))|
        (data_inflation_['event_lower'].str.contains('hicp'))|
        (data_inflation_['event_lower'].str.contains('inflation'))|
        (data_inflation_['event_lower'].str.contains('wage'))|
        (data_inflation_['event_lower'].str.contains('price index'))&
        ~(data_inflation_['event_lower'].str.contains('wpi'))&
        ~(data_inflation_['event_lower'].str.contains('wages'))&
        ~(data_inflation_['event_lower'].str.contains('dairy'))&
        ~(data_inflation_['event_lower'].str.contains('milk'))
        ]
    
    data_employ_manual = data_employ_[
        (data_employ_['importance'].isin(['medium', 'high']))|
       ~(data_employ_['event_lower'].str.contains('count'))&
       ~(data_employ_['event_lower'].str.contains('employee'))&
       ~(data_employ_['event_lower'].str.contains('wages'))
       ]
    
    data_cb_manual = data_cb_[
        (data_cb_['importance'].isin(['medium', 'high']))|
        (data_cb_['event_lower'].str.contains('speaks'))|
        (data_cb_['event_lower'].str.contains('decision'))|
        (data_cb_['event_lower'].str.contains('statement'))|
        (data_cb_['event_lower'].str.contains('minutes'))|
        (data_cb_['event_lower'].str.contains('meetings'))|
        (data_cb_['event_lower'].str.contains('press'))|
        (data_cb_['event_lower'].str.contains('conference'))|
        (data_cb_['event_lower'].str.contains('rate'))
        ]
    
    data_activity_manual = data_activity_[
        (data_activity_['importance'].isin(['medium', 'high']))|
        ~(data_activity_['event_lower'].str.contains('cpi'))&
        ~(data_activity_['event_lower'].str.contains('ppi'))&
        ~(data_activity_['event_lower'].str.contains('pce'))&
        ~(data_activity_['event_lower'].str.contains('hicp'))&
        ~(data_activity_['event_lower'].str.contains('inflation'))&
        ~(data_activity_['event_lower'].str.contains('wage'))&
        ~(data_activity_['event_lower'].str.contains('price index'))&
        ~(data_activity_['event_lower'].str.contains('wpi'))&        
        ~(data_activity_['event_lower'].str.contains('count'))&
        ~(data_activity_['event_lower'].str.contains('employee'))&
        ~(data_activity_['event_lower'].str.contains('speaks'))&
        ~(data_activity_['event_lower'].str.contains('decision'))&
        ~(data_activity_['event_lower'].str.contains('statement'))&
        ~(data_activity_['event_lower'].str.contains('minutes'))&
        ~(data_activity_['event_lower'].str.contains('meetings'))&
        ~(data_activity_['event_lower'].str.contains('press'))&
        ~(data_activity_['event_lower'].str.contains('arrivals'))&
        ~(data_activity_['event_lower'].str.contains('visitors'))&
        ~(data_activity_['event_lower'].str.contains('pcsi'))&
        ~(data_activity_['event_lower'].str.contains('capacity'))&
        ~(data_activity_['event_lower'].str.contains('tool'))&
        ~(data_activity_['event_lower'].str.contains('oil'))&
        ~(data_activity_['event_lower'].str.contains('gasoline'))&
        ~(data_activity_['event_lower'].str.contains('fuel'))&
        ~(data_activity_['event_lower'].str.contains('approvals'))&
        ~(data_activity_['event_lower'].str.contains('permits'))&
        ~(data_activity_['event_lower'].str.contains('migration'))&
        ~(data_activity_['event_lower'].str.contains('bonds'))&
        ~(data_activity_['event_lower'].str.contains('eia'))&
        ~(data_activity_['event_lower'].str.contains('financing'))&
        ~(data_activity_['event_lower'].str.contains('social'))&
        ~(data_activity_['event_lower'].str.contains('electronic'))&
        ~(data_activity_['event_lower'].str.contains('control'))&
        ~(data_activity_['event_lower'].str.contains('purchases'))&
        ~(data_activity_['event_lower'].str.contains('coincident'))&
        ~(data_activity_['event_lower'].str.contains('stocks'))&
        ~(data_activity_['event_lower'].str.contains('reuters'))&
        ~(data_activity_['event_lower'].str.contains('leading'))&
        ~(data_activity_['event_lower'].str.contains('wages'))&
        ~(data_activity_['event_lower'].str.contains('construction'))&
        ~(data_activity_['event_lower'].str.contains('dairy'))&
        ~(data_activity_['event_lower'].str.contains('milk'))&
        ~(data_activity_['event_lower'].str.contains('fgv'))&
        ~(data_activity_['event_lower'].str.contains('kc'))&
        ~(data_activity_['event_lower'].str.contains('machinery'))&
        ~(data_activity_['event_lower'].str.contains('foreign'))&
        ~(data_activity_['event_lower'].str.contains('inventories'))&
        ~(data_activity_['event_lower'].str.contains('climate'))&
        ~(data_activity_['event_lower'].str.contains('building'))&
        ~(data_activity_['event_lower'].str.contains('registration'))&
        ~(data_activity_['event_lower'].str.contains('profit'))&
        ~(data_activity_['event_lower'].str.contains('revenues'))&
        ~(data_activity_['event_lower'].str.contains('climate'))&
        ~(data_activity_['event_lower'].str.contains('debt-to-gdp'))&
        ~(data_activity_['event_lower'].str.contains('prime rate'))&
        ~(data_activity_['event_lower'].str.contains('halifax'))&
        ~(data_activity_['event_lower'].str.contains('rics '))&
        ~(data_activity_['event_lower'].str.contains('bsi large'))&
        ~(data_activity_['event_lower'].str.contains('eia'))&
        ~(data_activity_['event_lower'].str.contains('brc'))
        ]                                             
    
    #data_cb_dedupe = dedupe_top(data_cb_manual, threshold=0.0, top_n_words=5).drop('event_lower', axis=1)
    data_inflation_deduped = dedupe_top(data_inflation_manual, threshold=0.2, top_n_words=2).drop('event_lower', axis=1)
    data_employ_deduped = dedupe_top(data_employ_manual, threshold=0.0, top_n_words=2).drop('event_lower', axis=1)
    data_activity_deduped = dedupe_top(data_activity_manual, threshold=0.0, top_n_words=2).drop('event_lower', axis=1)   
    
    data = pd.concat([
        data_inflation_deduped.assign(category='inflation'),
        data_employ_deduped.assign(category='employ'),
        data_cb_manual.assign(category='cb'),
        data_activity_deduped.assign(category='activity'),
        data_holiday_.assign(category='holiday')]).drop('event_lower', axis=1).drop_duplicates()
    
    return data

#%% #SETUP
def get_and_filter_events(
    next_days=10,
    include_holidays=True,
    countries=None,
    categories=None,
    timezone='America/Sao_Paulo',
    verify_tls=None,
):
    #export_eco_data_global()
    eco_data = get_calendar_from_investing(
        next_days,
        include_holidays=include_holidays,
        countries=countries,
        categories=categories,
        timezone=timezone,
        verify_tls=verify_tls,
    ).sort_values(by='date')
    eco_data_fcast = eco_data.copy()
    #Remove Month Reference on event 
    eco_data_fcast = eco_data_fcast.drop_duplicates()
    eco_data_fcast_importance = eco_data_fcast[eco_data_fcast['importance'].isin(['high', 'medium', 'low'])]
    eco_data_fcast_importance['date'] = pd.to_datetime(eco_data_fcast_importance['date'], format='%d/%m/%Y').copy()
    eco_data_fcast_importance['datetime'] = eco_data_fcast_importance['date'].dt.strftime('%Y-%m-%d') + ' ' + eco_data_fcast_importance['time'].astype('string').copy()
    return eco_data_fcast_importance.drop(['date', 'time'], axis=1).sort_values(by=['datetime', 'zone'])

def export_events():
    """Compatibility name retained for callers of the former Excel flow."""
    return get_and_filter_events()


def logging_func():
    print('__name__', __name__)
    print('__file__', __file__)
    print('__doc__', __doc__)
    print('__package__', __package__)

def get_com_object(object_type='my_calendar', calendar_name=None):
    """Compatibility helper around the new Outlook adapter."""
    client = OutlookClient(calendar_name=calendar_name).connect()
    if object_type == 'outlook':
        return client.application
    if object_type == 'calendar':
        return client.namespace.GetDefaultFolder(9)
    if object_type == 'my_calendar':
        return client.calendar
    raise ValueError(f'object_type inválido: {object_type}')

def find_existing(subject_ls:list, start_dt:datetime, obj_type='my_calendar', calendar_name=None):
    """Retorna compromissos com mesmo assunto e data/hora exatas.

    Duplicatas não são apagadas automaticamente para evitar alterar eventos
    criados manualmente pelo usuário.
    """
    client = OutlookClient(calendar_name=calendar_name).connect()
    matched = []
    for subject in subject_ls:
        matched.extend(client.existing_appointments(subject, start_dt))
    unique = []
    for item in matched:
        if not any(item is existing for existing in unique):
            unique.append(item)
    return unique


ZONE_TO_CODE = {
        'United States': 'USD',
        'Euro Zone': 'EUR',
        'Brazil': 'BRL',
        'China': 'CNH', 
        'Australia': 'AUD',
        'Germany': 'GER',
        'Hungary': 'HUF', 
        'Czech Republic': 'CZK',
        'United Kingdom': 'GBP',
        'Canada':'CAD', 
        'Japan': 'JPY',
        'New Zealand': 'NZD',
        'Poland': 'PLN',
        'South Africa': 'ZAR', 
        'Mexico': 'MXN',
        'Norway': 'NOK',
        'South Korea': 'KRW', 
        'Sweden':'SEK', 
        'Switzerland': 'CHF',
        'France':'FRA',
        'India': 'INR',
        'Türkiye':'TRY',
        'Turkey':'TRY',
        'Spain': 'SPN',
        'Italy': 'ITL',
        'Chile': 'CLP', 
        'Colombia': 'COP'}


def _format_subject(row):
    region = str(row['zone']).replace('  ', ' ')
    zone = ' '.join(map(str.capitalize, region.split())).strip()
    zone_code = ZONE_TO_CODE.get(zone, zone)
    return f"{zone_code} | {row['event']}"


def sync_events(df, config=None, dry_run=False, client=None):
    """Create or update events in Outlook, preserving the existing pipeline."""
    config = config or AppConfig()
    config.validate()
    outlook = client or OutlookClient(config.calendar_name).connect()
    counts = {'created': 0, 'updated': 0, 'duplicates': 0}

    for idx, row in df.iterrows():
        start = pd.to_datetime(row['datetime'], errors='coerce')
        if pd.isna(start):
            raise ValueError(f"Evento {idx} possui datetime inválido: {row['datetime']!r}")

        subject = _format_subject(row)
        category = f"MacroCalendar - {row.get('category', 'event')}"
        try:
            action, has_duplicates = outlook.upsert_event(
                subject=subject,
                start=start.to_pydatetime() if hasattr(start, 'to_pydatetime') else start,
                duration_minutes=config.duration_minutes,
                reminder_minutes=config.reminder_minutes,
                category=category,
                dry_run=dry_run,
            )
        except Exception as exc:
            raise RuntimeError(f"Falha ao sincronizar '{subject}': {exc}") from exc
        counts['updated' if action == 'atualizar' else 'created'] += 1
        if has_duplicates:
            counts['duplicates'] += 1
        suffix = ' (dry-run)' if dry_run else ''
        duplicate_note = ' [mais de uma correspondência]' if has_duplicates else ''
        display_datetime = start.strftime('%d/%m/%Y %H:%M')
        print(f"  {action.capitalize()}: {display_datetime} | {subject}{suffix}{duplicate_note}")

    return counts


def add_to_agenda(df=None, calendar_name=None, dry_run=False, config=None):
    """Backward-compatible entry point for the CLI and existing scripts."""
    config = config or AppConfig()
    if calendar_name is not None:
        config.calendar_name = calendar_name
    if df is None:
        df = get_and_filter_events(
            config.days,
            include_holidays=config.include_holidays,
            countries=config.countries,
            categories=config.categories,
            timezone=config.timezone,
            verify_tls=config.verify_tls,
        )
    return sync_events(df, config=config, dry_run=dry_run)

def _run_scheduled_task():
    """Funções que só devem rodar no Task Scheduler"""
    print('Running AS TASK!!!')
    add_to_agenda()

#print('__name__', __name__)

# <-- aqui o sentinela
if __name__ == "__main__":
    print('__name__', __name__)
    _run_scheduled_task()
