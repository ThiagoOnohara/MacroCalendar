# -*- coding: utf-8 -*-
"""
@author: thiago.onohara
"""

import investpy
from datetime import date, timedelta,datetime
import pandas as pd
from abc import ABC
import os

investpy.user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.5845.111 Safari/537.3'

class USER(ABC):
    email = 'your_email'
    name = os.getlogin()

'''Getting Next Inflation Release Dates'''

class EcoCalendar:
    '''
   Economic Categories:
   'credit',
   'inflation'
   'employment'
   'activity'
   'central_banks'
   'balance'
   'bonds'
   
   Country (str format: lower case)
   EXAMPLES
   - united states
   - china
   - brazil
   - hungary
   - mexico
   - south africa
   
   To see available categories:
       EcoCalendar.available_categories
    '''
    def __init__(
            self,
            dt_format='%d/%m/%Y'):
        
        self.available_categories = [
            'credit',
            'inflation',
            'employment',
            'activity',
            'central_banks',
            'balance', 
            'bonds']
        
        self.dt_format = dt_format
        
    def get_economic_calendar(
            self,
            from_date:datetime,
            to_date:datetime,
            categories:list=None,
            countries:list=None) -> pd.DataFrame:
        
        
        data = investpy.economic_calendar(
            from_date=from_date.strftime(self.dt_format),
            to_date=to_date.strftime(self.dt_format),
            categories=categories,
            countries=countries)
        
        if data.empty:
            raise Exception(f'''Dataframe Returns Empty 
                            Request {USER.name}''')
        else:
            print('Calendar obtained to: ', data.dropna(subset='currency')['currency'].unique())
            print('Calendar obtained to: ', data.dropna(subset='currency')['zone'].unique())
            return data.dropna(subset='currency')
            
    

