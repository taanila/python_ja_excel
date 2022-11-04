# Tuoreimman datan nouto ja siivous
# Ennuste 36 kuukautta eteenpäin
# Grafiikka, jossa aikasarja ja liukuva 12 kuukauden keskiarvo
# Grafiikka, jossa ennuste 36 kuukautta eteenpäin
# Data, ennusteet ja grafiikat Exceliin

# Toimivuus testattu Anacondan versiolla Anaconda3-2021.11-Windows-x86_64, johon asennettu lisäksi
# pandas-datareader versio 0.10.0 ja xlwings versio 0.24.9

# Tämän voi suorittaa komentoriviltä komennolla: python co2.py


import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
from statsmodels.tsa.api import ExponentialSmoothing
import warnings
warnings.filterwarnings('ignore')
plt.style.use('seaborn-whitegrid')

# Datan nouto ja siivous
df = pd.read_csv('https://www.esrl.noaa.gov/gmd/webdata/ccgg/trends/co2/co2_mm_mlo.txt',\
sep='\s+', skiprows=54, usecols=[0, 1, 3], names=['year', 'month', 'average'])
df.index = pd.to_datetime(df['year'].astype(str) + df['month'].astype(str), format='%Y%m')
df = df.drop(['year', 'month'], axis=1)

# Ennustemalli kolminkertaista eksponentiaalista tasoitusta käyttäen
malli = ExponentialSmoothing(df['average'], trend='add', seasonal='mul', seasonal_periods=12, freq='MS').fit()

# Ennusteet 36 kuukautta eteenpäin
next_date = df.index.to_series().iloc[-1]
index = pd.date_range(f'{next_date.year}-{next_date.month+1}-{next_date.day}', periods=36, freq='MS')
ennusteet = malli.forecast(36)
df_ennuste = pd.DataFrame(data=ennusteet, index=index, columns=['Ennuste'])

# Grafiikka
fig, ax = plt.subplots(nrows=2, ncols=1, figsize=(8, 8), tight_layout=True)
df.plot(ax=ax[0], title='Hiilidioksidipitoisuudet ja 12 kuukauden liukuva keskiarvo', legend=False)
df.rolling(12).mean().plot(ax=ax[0], legend=False)
df['2000':]['average'].plot(ax=ax[1], title = 'Ennuste 3 vuotta eteenpäin')
df_ennuste['Ennuste'].plot(ax=ax[1])

# Excel-työkirjan alustus
wb = xw.Book()
sht = wb.sheets[0]
sht.name = 'co2'

# Data ja ennusteet Exceliin
xw.Range('A1:H1').add_hyperlink(address='https://www.esrl.noaa.gov/gmd/webdata/ccgg/trends/co2/co2_mm_mlo.txt', \
    text_to_display='Lähde: https://www.esrl.noaa.gov/gmd/webdata/ccgg/trends/co2/co2_mm_mlo.txt')
xw.Range('A3').value = df
xw.Range('B3').value = 'Monthly average ppm'
xw.Range('D43').value = df_ennuste

# Grafiikka Exceliin
sht.pictures.add(fig, anchor=xw.Range('D3'))