from calendar import month
import requests
import bs4
from urllib.request import urlretrieve
import pandas as pd
from xlsxwriter import Workbook
import datetime


def main():
    pd.options.mode.chained_assignment = None # Disables annoying and useless warning
    #get_data()
    caged_to_excel()
    #writer = caged_to_excel()
    #wb = writer.book

    #writer.save()


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
"""""""""

def caged_to_excel():
    # Gets old data
    old_df = pd.read_excel('Tabela velho caged.xls', sheet_name='tabela10.1', header=5)
    old_balance, old_dates = old_df['Total das Atividades'].drop([84, 85]), old_df['Mês/ Ano'].drop([84, 85])
    old_df = pd.concat([old_dates, old_balance], axis=1)
    old_df = old_df.rename(columns={'Total das Atividades': 'Saldos', 'Mês/ Ano': 'Mês'})
    print(old_df)

    # Gets newer data
    new_df = pd.read_excel('tabela caged.xlsx', sheet_name='Tabela 5.1', header=4)
    new_balance, new_dates = new_df['Saldos'], new_df['Mês']
    new_df = pd.concat([new_dates, new_balance], axis=1)

    # Merges them into the same DataFrame
    df = pd.concat([old_df, new_df], axis=0)
    writer = pd.ExcelWriter('Saldo caged.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Dados')
    writer.save()
"""""""""
def caged_to_excel():
    # Gets old data
    old_df = pd.read_excel('Tabela velho caged.xls', sheet_name='tabela10.1', header=5)
    old_balance, old_dates = old_df['Total das Atividades'].drop([84, 85]), old_df['Mês/ Ano'].drop([84, 85])
    old_dates = old_dates.to_list()

    # Gets newer data
    new_df = pd.read_excel('tabela caged.xlsx', sheet_name='Tabela 5.1', header=4)
    new_balance, new_dates = new_df['Saldos'], new_df['Mês']
    new_dates = date_converter(new_dates.to_list())

    # Merges them into the same DataFrame
    dates = old_dates + new_dates
    df = pd.concat([old_balance, new_balance], axis=0)
    print(df)


# Converts list of date strings into list of datetime objects to write into file more easily 
def date_converter(list):
    for item in list:
        try:
            item = datetime.datetime.fromisoformat(month_converter(item))

        except:
            item = 0

    return list


# Converts months to ISO 8601 dates    
def month_converter(date):
    months = {'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8, 'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12}
    
    try:
        month, year = date.split("/")
        month = months[month]
        return f'{year}-{month:02d}-01'

    except:
        return 0    


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

