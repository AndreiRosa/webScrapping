import phantomjs
from selenium import webdriver   # for webdriver
from selenium.webdriver.support.ui import WebDriverWait     # For implicit and explict waits
from selenium.webdriver.chrome.options import Options       # For suppressing the browser
from bs4 import BeautifulSoup
import webbrowser
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import re
import numpy as np
from openpyxl import Workbook
import csv

def clean_string(text_string):
    index_start = text_string.find('>')
    index_end = text_string.find('/')
    return text_string[index_start+1:index_end-1]

def separate_range_from_estimate(text_string):
    list_string = text_string.split('*<p><small>')
    estimate = list_string[0][3:]
    range = list_string[1][4:-2]
    range = range.replace("R$ ", '')

    return estimate, range

def Search_apartment(apto):
    try:
        option = webdriver.ChromeOptions()
        option.add_argument('headless')
        driver = webdriver.Chrome('C://Users//5927951//Downloads//chromedriver_win32//chromedriver.exe', options=option)
        driver.get('https://123i.uol.com.br/#')
        button = driver.find_element_by_id('search_goal_estimate')  # HTML ID
        button.click()

        search_input_box = driver.find_element_by_name('q')
        search_input_box.send_keys(apto)
        search_input_box.send_keys(u'\ue007')

        content = driver.page_source
        soup = BeautifulSoup(content, 'html.parser')

        squared_meter = soup.findAll('div', attrs={'class': 'area_useful'})
        price = soup.findAll('div', attrs={'class': 'bird_estimate_average'})


        squared_meter = clean_string(str(squared_meter[1]))
        squared_meter = squared_meter[:-2]
        price = clean_string(str(price[1]))
        estimate, range = separate_range_from_estimate(price)
        link = driver.current_url
        #else:
    except:
            squared_meter, estimate, range, link = 'NULL', 'NULL', 'NULL', 'NULL'

    return squared_meter, estimate, range, link

def get_data_web(filename):
    df = pd.read_excel(filename, sheetname='Sheet1')
    addr = df['ENDEREÇO']

    new_df = pd.DataFrame(columns=['Código do Imóvel', 'Endereço', 'Metragem Quadrada da Unidade (m2)',
                                   'Range (R$)', 'Estimativa Confiável(R$)', 'Link'])
    for i, a in enumerate(addr):
        if  pd.isnull(a) == False:
            a = re.sub('[-/,.]', ' ', a)  # Remove special characters
            match = re.search('\s*\d+\s*', a)
            last_number_index = match.end()
            search_string = a[:last_number_index]
            sq, estimate, range, link = Search_apartment(search_string)
            print(df['CÓDIGO DO IMÓVEL'][i])
            new_df = new_df.append({'Código do Imóvel': df['CÓDIGO DO IMÓVEL'][i], 'Endereço' : df['ENDEREÇO'][i],
                                    'Metragem Quadrada da Unidade (m2)': sq, 'Range (R$)': range,
                                    'Estimativa Confiável(R$)': estimate, 'Link': link}, ignore_index=True)

    new_df.to_csv('export_dataframe.csv', index=None, header=True)

filename = 'Base de apartamentos.xlsx'
get_data_web(filename)

