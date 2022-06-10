from calendar import month
from xlsxwriter import Workbook
from urllib.request import urlretrieve
from data import Data
from math import isnan
import requests
import bs4
import pandas as pd


total_entries = None


def main():
    pd.options.mode.chained_assignment = None # Disables annoying and useless warning
    #get_data()
    workbook, worksheet = caged_to_excel()
    write_formulas(workbook, worksheet)
    #writer = caged_to_excel()
    #wb = writer.book

    workbook.close()


# Writes formulas whose values will later be used to make the chart
def write_formulas(workbook, worksheet):
    global total_entries
    
    # Writes headers
    worksheet.write('C1', 'Acumulado 12 meses')
    worksheet.write('D1', 'Saldo mil')
    worksheet.write('E1', 'Acumulado 12 meses mil')
    
    # Writes formulas
    number_format = workbook.add_format({'num_format': '#,##0.0'})
    for i in range(total_entries):
        worksheet.write_formula(f'C{i + 13}', f'=SUM(B{i + 2}:B{i + 13})')
        worksheet.write_formula(f'D{i + 2}', f'=B{i + 2}/1000', number_format)
        worksheet.write_formula(f'E{i + 2}', f'=C{i + 2}/1000', number_format)


def make_chart(wb):
    chartsheet = wb.add_chartsheet('Gráfico')

    chart = wb.add_chart({'type': 'line'})
    chart.set_x_axis({'values': '=Dados!$A$2:A$29', 'date_axis': True, 'label_position': 'low'})
    chart.add_series({'values': '=Dados!$B$2:B$29'})
    chartsheet.set_chart(chart)
    
    
# Extracts the necessary data from the caged sheet. Returns workbook and worksheet.    
def caged_to_excel():
    global total_entries
    
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

    # Saves global variable
    total_entries = len(entries)

    # Writes into Excel file
    workbook = Workbook('Saldo.xlsx')
    worksheet = workbook.add_worksheet('Dados')

    # Writes headers
    worksheet.write('A1', 'Mês')
    worksheet.write('B1', 'Saldo')

    # Writes data
    date_format = workbook.add_format({'num_format': 'mmm-yy'})
    for i in range(total_entries):
        worksheet.write_datetime(i + 1, 0, entries[i].date, date_format)
        worksheet.write(i + 1, 1, entries[i].value)

    return workbook, worksheet
        

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

