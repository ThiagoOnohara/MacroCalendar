# -*- coding: utf-8 -*-
"""
Created on Fri Nov  8 12:39:11 2024

@author: thiago.onohara
"""

import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import sys
from functools import reduce
from datetime import datetime, timedelta
import re
sys.path.append('F:/Front/Moedas/Base/')
sys.path.append('G:/Front/Moedas/Utils/Thiago/scripts/')
from eco_calendar import EcoCalendar
from investing_calendar import InvestingAPIClient
import xlwings as xw
import win32com.client

'''
try:
    eco_calendar = EcoCalendar()
    countries=[
        'spain',
        'france',
        'united kingdom',
        'italy',
        'germany',
        'united states', 
        'japan',
        'brazil',
        'mexico'
    ]
    
    test = eco_calendar.get_economic_calendar(
        from_date=datetime.today()-timedelta(days=10),
        to_date=datetime.today()+timedelta(days=15),
        categories=['economic_activity'],
        countries=countries)
    
    if not test.empty:
        calendar = eco_calendar
        
except:'''

calendar = InvestingAPIClient()

#%% UTILS
#Metrics (não ativas)
sharpe = lambda series: np.sqrt(252)*series.mean()/series.std()
hitratio = lambda series: series[series>0].count()/series.count()
pct = lambda series: series[series!=0].dropna().count()/series[series!=0].dropna().shape[0]
count = lambda series: series.count()

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

    result = df.groupby('group_key', group_keys=False).apply(filter_group)
    return result.drop(columns='group_key').reset_index(drop=True).sort_values(by=['date', 'time', 'zone'])

def get_calendar_from_investing(next_days=7):
    countries_inflation =list(country_total.keys())
    countries_central_banks =list(country_cb.keys())
    countries_employment =list(country_employment.keys())
    countries_activity =list(country_total.keys())

    #Inflation Events
    data_inflation = calendar.get_economic_calendar(
        from_date=datetime.today(),
        to_date=datetime.today()+timedelta(days=next_days),
        categories=[
         'inflation'],
        countries=countries_inflation)
    #Employment Events
    data_employ = calendar.get_economic_calendar(
        from_date=datetime.today(),
        to_date=datetime.today()+timedelta(days=next_days),
        categories=['employment'],
        countries=countries_employment)
    #Central Banks
    data_cb = calendar.get_economic_calendar(
        from_date=datetime.today(),
        to_date=datetime.today()+timedelta(days=next_days),
        categories=['central_banks'],
        countries=countries_central_banks)
    #Activity
    data_activity = calendar.get_economic_calendar(
        from_date=datetime.today(),
        to_date=datetime.today()+timedelta(days=next_days),
        categories=['economic_activity'],
        countries=countries_activity)
    
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
        data_activity_deduped.assign(category='activity')]).drop('event_lower', axis=1).drop_duplicates()
    
    return data

#%% #SETUP
def get_and_filter_events(next_days=10):
    #export_eco_data_global()
    eco_data = get_calendar_from_investing(next_days).sort_values(by='date')
    eco_data_fcast = eco_data.copy()
    #Remove Month Reference on event 
    eco_data_fcast = eco_data_fcast.drop_duplicates()
    eco_data_fcast_importance = eco_data_fcast[eco_data_fcast['importance'].isin(['high', 'medium', 'low'])]
    eco_data_fcast_importance['date'] = pd.to_datetime(eco_data_fcast_importance['date'], format='%d/%m/%Y').copy()
    eco_data_fcast_importance['datetime'] = eco_data_fcast_importance['date'].dt.strftime('%Y-%m-%d') + ' ' + eco_data_fcast_importance['time'].astype('string').copy()
    return eco_data_fcast_importance.drop(['date', 'time'], axis=1).sort_values(by=['datetime', 'zone'])

#%% xw funcs
@xw.sub
def export_events():
    events_df = get_and_filter_events()
    
    try:
        wb = xw.Book.caller()
        events_sh = wb.sheets('EVENTS')
        events_sh.range('A5').expand().clear()    
        events_sh.range('A5').value = events_df 
    except:
        return events_df
    
@xw.sub
def logging_func():
    print('__name__', __name__)
    print('__file__', __file__)
    print('__doc__', __doc__)
    print('__package__', __package__)

def get_com_object(object_type='my_calendar'):    
    outlook = win32com.client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    if object_type == 'outlook':
        return outlook
    if object_type == 'calendar':
        return namespace.GetDefaultFolder(9) #9=olFolderCalendar
    if object_type == 'my_calendar':
        return namespace.GetDefaultFolder(9).Folders['FX'] # nome exato da subpasta

def find_existing(subject_ls:list, start_dt:datetime, obj_type='my_calendar'):
    """Retorna uma lista de AppointmentItems com mesmo assunto e data exata."""
    items = get_com_object(object_type=obj_type).Items
    # Outlook guarda Start como string “YYYY-MM-DD HH:MM”
    # precisamos filtrar pela data exata (pode incluir hora, se quiser)
    matched = [item for item in items
               if item.Subject in subject_ls 
               and item.Start.strftime('%Y-%m-%d') == start_dt.strftime('%Y-%m-%d')]
    
    if len(matched) > 1:
        print('Removendo Itens Duplicados: ', len(matched))
        duplicated_items = matched[1:]
        for duplicated_item in duplicated_items:
            print('Removendo: ', duplicated_item.Subject, duplicated_item.Start.strftime('%Y-%m-%d %H:%M'))
            duplicated_item.Delete()
    return list(matched[:1])  # pode retornar 0 ou mais


@xw.sub
def add_to_agenda():
    mail_adress = '@legacycapital.com.br'
    emails = ['thiago.onohara']
    emails_completos = [i+mail_adress for i in emails]
    # Suponha df com colunas: date (YYYY-MM-DD), time (HH:MM), zone (e.g. 'America/Sao_Paulo'), event (str)
    
    zone_to_code = {
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
    
    try:
        wb = xw.Book.caller()
        events_sh = wb.sheets('EVENTS')
        range_address = 'A5:K1000'
        df = events_sh.range(range_address).options(pd.DataFrame, index=False).value.dropna(axis=0, how='all')
        print('Eventos na Planilha:', df, sep='\n')
        task = False
    except:
        df = get_and_filter_events()
        task = True
        
    #%% Export Eventos to Sheet Events

    for idx, row in df.iterrows():

        # formata datetime com timezone
        region = row['zone'].replace('  ', ' ')
        zone = ' '.join(map(str.capitalize, region.split())).strip()
        category = row['category']
        start = row['datetime'] if not task else datetime.strptime(row['datetime'], '%Y-%m-%d %H:%M')
        print('start', start)
        
        remind_before = 15 #minutes  # por exemplo, duração 15min
        evento = row['event']
         # Monta o datetime com timezone
        end = start + pd.Timedelta(minutes=30)
    
        subject = f"{zone} | {evento}"
        subject_zone_code = f"{zone_to_code.get(zone)} | {evento}"

        exists = find_existing([subject, subject_zone_code], start)
        if exists:
            # atualiza o primeiro que encontrar
            appt = exists[0]
            print(f"→ Atualizando evento existente: {subject}")
        else:
            # cria novo
            appt = get_com_object(object_type='my_calendar').Items.Add(1)
            print(f"→ Criando novo evento: {subject}")
        
        appt.Subject = subject_zone_code
        appt.Start = start.strftime('%Y-%m-%d %H:%M')
        appt.End = end.strftime('%Y-%m-%d %H:%M')
    
        # Deixa como "Free" no calendário, mas com lembrete
        appt.BusyStatus = 0             # 0 = olFree
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = remind_before
        
        def category_to_color(category:str):
            
            category_color_meta = {
                'cb':'Dark', 
                'inflation':'Green',
                'activity': 'Blue',
                'employ': 'Red'}
            
            return category_color_meta.get(category)
        
        color = category_to_color(category)
        appt.Categories = f"{color} category"
    
        # marca os Required Attendees, mas NÃO converte em Meeting    
        # Aqui: transforma lista em string "email1; email2; ..."
        if isinstance(emails_completos, list) and emails_completos:
            appt.RequiredAttendees = "; ".join(emails_completos)
        else:
            appt.RequiredAttendees = ""
    
        # só salva—não envia convites
        appt.Save()
        print(f"✔ Appointment '{subject}' às {start} (free) criado com lembrete.")
        
def send_logging_email():
        
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'thiago.onohara@legacycapital.com.br'
    mail.Subject = 'TASK COMPLETA -> Calendário Atualizado'
    mail.Send()

#%%
@xw.sub
def month_end_rates_task():
    
    mail_adress = '@legacycapital.com.br'
    emails = ['thiago.onohara', 'mleal']
    emails_completos = [i+mail_adress for i in emails]
        
    today = pd.to_datetime('today').date()
    end_of_year = pd.to_datetime(f'{today.year+1}-12-31').date()
    dates_rng = pd.date_range(start=today, end=end_of_year, freq='B', name='dates')
    dates_rng_df = dates_rng.to_frame(index=False)
    dates_rng_df['month'] = dates_rng_df['dates'].dt.month
    dates_rng_df['Ym'] = dates_rng_df['dates'].dt.strftime('%Y-%m')

    last_dt_month = dates_rng_df.groupby('Ym')['dates'].max()
    last_dt_month = last_dt_month[last_dt_month.dt.day>=25]
    
    month_end_dates = last_dt_month.to_frame('month_end')
    month_end_dates['trade_start'] = month_end_dates['month_end'] - pd.tseries.offsets.BusinessDay(4)
    month_end_dates['trade_end'] = month_end_dates['month_end'] + pd.tseries.offsets.BusinessDay(1)
    
    month_end_dates_ = month_end_dates.reset_index()
    
    for i, row in month_end_dates_.iterrows():
        remind_before = 15 #minutes  # por exemplo, duração 15min
        evento = 'Month End Rates'
        
        month_end = row['month_end']
        month_ref = month_end.strftime('%Y-%m')
        
        trade_start = row['trade_start']
        trade_end = row['trade_end']

        
        start = datetime(year=trade_start.year, month=trade_start.month, day=trade_start.day, hour=8)
         # Monta o datetime com timezone
        end = datetime(year=trade_end.year, month=trade_end.month, day=trade_end.day, hour=8)
    
        subject = f"{month_ref} | {evento}"
        
        # Cria o Appointment (1 = olAppointmentItem)
        appt = get_com_object(object_type='outlook').CreateItem(1)
        
        exists = find_existing([subject], start)
        
        if exists:
            # atualiza o primeiro que encontrar
            appt = exists[0]
            print(f"→ Atualizando evento existente: {subject}")
        else:
            # cria novo
            appt = get_com_object(object_type='my_calendar').Items.Add(1)
            print(f"→ Criando novo evento: {subject}")
        
        appt.Subject = subject
        appt.Start = start.strftime('%Y-%m-%d %H:%M')
        appt.End = end.strftime('%Y-%m-%d %H:%M')
    
        # Deixa como "Free" no calendário, mas com lembrete
        appt.BusyStatus = 0             # 0 = olFree
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = remind_before
        
        appt.Categories = "Yellow category"
    
        # marca os Required Attendees, mas NÃO converte em Meeting    
        # Aqui: transforma lista em string "email1; email2; ..."
        if isinstance(emails_completos, list) and emails_completos:
            appt.RequiredAttendees = "; ".join(emails_completos)
        else:
            appt.RequiredAttendees = ""
    
        # só salva—não envia convites
        appt.Save()
        print(f"✔ Appointment '{subject}' às {start} (free) criado com lembrete.")
            
        
@xw.sub
def qra_announcements_task():
    
    mail_adress = '@legacycapital.com.br'
    emails = ['thiago.onohara', 'mleal', 'marco.lyrio']
    emails_completos = [i+mail_adress for i in emails]
        
    today = pd.to_datetime('today').date()
    start_of_year = pd.to_datetime(f'{today.year}-01-01').date()
    end_of_next_year = pd.to_datetime(f'{today.year+1}-12-31').date()
    dates_rng = pd.date_range(start=start_of_year, end=end_of_next_year , freq='B', name='dates')

    months_final = [1, 4, 7, 10] #fim dos trimestres
    months_start = [2, 5, 8, 11] #se passar o fim, primeira semana após
    
    dates_rng_mon = dates_rng[dates_rng.weekday==0]
    
    dates_rng_mon_df = dates_rng_mon.to_frame(index=False)
    
    dates_rng_mon_df_final = dates_rng_mon_df[dates_rng_mon_df['dates'].dt.month.isin(months_final)]
    dates_rng_mon_df_start = dates_rng_mon_df[dates_rng_mon_df['dates'].dt.month.isin(months_start)]
    
    dates_rng_mon_df_final_last = dates_rng_mon_df_final.groupby(dates_rng_mon_df_final['dates'].dt.month).last()
    dates_rng_mon_df_start_first = dates_rng_mon_df_start.groupby(dates_rng_mon_df_start['dates'].dt.month).first()
    
    possible_year_qra_dates = pd.concat([dates_rng_mon_df_final_last, dates_rng_mon_df_start_first])
    possible_year_qra_dates.index.name = 'month'
    
    possible_year_qra_dates_ = possible_year_qra_dates.reset_index()
    
    possible_year_qra_dates_['trade_start'] = possible_year_qra_dates_['dates'] - pd.tseries.offsets.BusinessDay(3)
    possible_year_qra_dates_['trade_end'] = possible_year_qra_dates_['dates'] + pd.tseries.offsets.BusinessDay(3)
    possible_year_qra_dates_
    
    for i, row in possible_year_qra_dates_.iterrows():
        remind_before = 15  #minutes  # por exemplo, duração 15min
        evento = 'Possible QRA pré Drift [APLICAR TSY no FECHAMENTO], melhor retorno [Sexta-Feira: D-1 QRA]'
        
        qra = row['dates']
        quarter_ref = (qra.month - 1) // 3 + 1

        trade_start = row['trade_start']
        trade_end = row['trade_end']

        start = datetime(year=trade_start.year, month=trade_start.month, day=trade_start.day, hour=8)
        end = datetime(year=trade_end.year, month=trade_end.month, day=trade_end.day, hour=8)
    
        subject = f"Q{quarter_ref} | {evento}"
        
        # Cria o Appointment (1 = olAppointmentItem)
        appt = get_com_object(object_type='outlook').CreateItem(1)
        
        exists = find_existing([subject], start)
        
        if exists:
            # atualiza o primeiro que encontrar
            appt = exists[0]
            print(f"→ Atualizando evento existente: {subject}")
        else:
            # cria novo
            appt = get_com_object(object_type='my_calendar').Items.Add(1)
            print(f"→ Criando novo evento: {subject}")
        
        appt.Subject = subject
        appt.Start = start.strftime('%Y-%m-%d %H:%M')
        appt.End = end.strftime('%Y-%m-%d %H:%M')
    
        # Deixa como "Free" no calendário, mas com lembrete
        appt.BusyStatus = 0             # 0 = olFree
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = remind_before
        
        appt.Categories = "Orange category"
    
        # marca os Required Attendees, mas NÃO converte em Meeting    
        # Aqui: transforma lista em string "email1; email2; ..."
        if isinstance(emails_completos, list) and emails_completos:
            appt.RequiredAttendees = "; ".join(emails_completos)
        else:
            appt.RequiredAttendees = ""
    
        # só salva—não envia convites
        appt.Save()
        print(f"✔ Appointment '{subject}' às {start} (free) criado com lembrete.")
        
#%% For task execution>

def _run_scheduled_task():
    """Funções que só devem rodar no Task Scheduler"""
    print('Running AS TASK!!!')
    add_to_agenda()
    month_end_rates_task()
    qra_announcements_task()
    send_logging_email()

#print('__name__', __name__)

# <-- aqui o sentinela
if __name__ == "__main__":
    print('__name__', __name__)
    _run_scheduled_task()