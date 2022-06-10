from calendar import month
from xlsxwriter import Workbook
from urllib.request import urlretrieve
from data import Data
from math import isnan
import requests
import bs4
import pandas as pd
import datetime


def main():
    pd.options.mode.chained_assignment = None # Disables annoying and useless warning
    #get_data()
    workbook = caged_to_excel()
    #writer = caged_to_excel()
    #wb = writer.book

    workbook.close()


# Arranges sheet to later make the chart
def arrange_sheet(ws):
    ...


def make_chart(wb):
    chartsheet = wb.add_chartsheet('Gráfico')

    chart = wb.add_chart({'type': 'line'})
    chart.set_x_axis({'values': '=Dados!$A$2:A$29', 'date_axis': True, 'label_position': 'low'})
    chart.add_series({'values': '=Dados!$B$2:B$29'})
    chartsheet.set_chart(chart)
    
    
# Extracts the necessary data from the caged sheet. Returns workbook.    
def caged_to_excel():
    # Gets old data as list
    old_df = pd.read_excel('Tabela velho caged.xls', sheet_name='tabela10.1', header=5)
    old_balance, old_dates = old_df['Total das Atividades'].drop([84, 85]), old_df['Mês/ Ano'].drop([84, 85])
    old_balance, old_dates = old_balance.to_list(), old_dates.to_list()

    # Gets newer data as list
    new_df = pd.read_excel('tabela caged.xlsx', sheet_name='Tabela 5.1', header=4)
    new_balance, new_dates = new_df['Saldos'], new_df['Mês']
    new_balance, new_dates = new_balance.to_list(), new_dates.to_list()

    # Merges them into the same list
    balance = old_balance + new_balance
    dates = old_dates + new_dates
    entries = []

    for i in range(len(balance)):
        
        try:
            if isnan(dates[i]) or isnan(balance[i]):
                break
        except TypeError:
            pass

        entries.append(Data(dates[i], balance[i]))

    # Writes into Excel file
    workbook = Workbook('Saldo.xlsx')
    worksheet = workbook.add_worksheet('Dados')

    # Writes headers
    worksheet.write('A1', 'Mês')
    worksheet.write('B1', 'Saldo')

    # Writes data
    date_format = workbook.add_format({'num_format': 'mmm-yy'})
    for i in range(len(entries)):
        worksheet.write_datetime(i + 1, 0, entries[i].date, date_format)
        worksheet.write(i + 1, 1, entries[i].value)

    return workbook
        

# Gets the Excel files from the CAGED website
def get_data():    
    # Gets data prior to 2020
    old_url = 'http://pdet.mte.gov.br/images/ftp//dezembro2019/nacionais/4-tabelas.xls'
    urlretrieve(old_url, 'Tabela velho caged.xls')

    # Gets data for 2020 onwards
    new_url = requests.get('http://pdet.mte.gov.br/novo-caged?view=default')
    new_caged = bs4.BeautifulSoup(new_url.text, 'html.parser')
    new_link = new_caged.select('#content-section > div.row-fluid > div > div.row-fluid.module > div.listaservico.span8.module.span6 > ul > li.item-6057 > a')
    new_link = new_link[0].get('href')
    new_link = f'http://pdet.mte.gov.br{new_link}'
    urlretrieve(new_link, 'Tabela caged.xlsx')
    print("Successfully downloaded file")


if __name__ == '__main__':
    main()

