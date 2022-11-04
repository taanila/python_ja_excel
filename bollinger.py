import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import pandas_datareader.data as web
import sys

def bollinger1(stock, span=20):
    
    '''Piirtää Bollingerin nauhat Exceliin'''
    
    stock['MA'] = stock['Close'].rolling(span).mean()
    stock['STD'] = stock['Close'].rolling(span).std(ddof = 0) 
    stock['Upper'] = stock['MA'] + (stock['STD'] * 2)
    stock['Lower'] = stock['MA'] - (stock['STD'] * 2)

    fig, ax = plt.subplots()
    stock['Close'].plot(label = 'Close', color = 'black', figsize = (10, 6))
    stock['Upper'].plot(label = 'Upper', linestyle = '--', linewidth = 1, color = 'red')
    stock['MA'].plot(label = 'Middle', linestyle = '--', linewidth = 1.2, color = 'grey')
    stock['Lower'].plot(label = 'Lower', linestyle = '--', linewidth = 1, color = 'red')
    plt.gca().fill_between(stock.index, stock['Lower'], stock['Upper'], facecolor = 'yellow', alpha = 0.1)
    plt.legend()
    plt.title('BOLLINGER BANDS')

    return stock, fig


if __name__=='__main__':

    if (len(sys.argv) > 1):
        x = sys.argv[1]
    else:
        x = input('Anna osakkeen Yahoo Finance tunnus: ')

    try:
        stock = web.DataReader(x, start = '2021-1-1', data_source = 'yahoo')
    except:
        print('Tunnuksella ei löytynyt tietoja')
        exit()
    
    stock, fig = bollinger1(stock)

    wb = xw.Book()
    sht = wb.sheets[0]
    sht.name = 'bollinger'
    xw.Range('A1').value = x
    xw.Range('A30').value = stock
    sht.pictures.add(fig, anchor = xw.Range('A3'))