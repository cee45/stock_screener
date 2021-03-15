import requests, time, re, os
import pickle as pkl
from configparser import ConfigParser
import pandas as pd
from openpyxl import load_workbook


config = ConfigParser()
config.read('auth.ini')
td_key = config.get('auth', 'td_key')

url = 'https://api.tdameritrade.com/v1/instruments'

df = pd.read_excel('us_equities_companies.xlsx')
symbols = df['Symbol'].values.tolist()

start = 0
end = 500
files = []
while start < len(symbols): 
    tickers = symbols[start:end]

    payload = {'apikey' : td_key,
            'symbol': tickers,
            'projection':'fundamental'}

    results = requests.get(url,params=payload)
    data =  results.json()
    f_name = time.asctime() + '.pkl'
    f_name = re.sub('[ :]', '_', f_name)
    files.append(f_name)
    with open(f_name,'wb') as file:
        pkl.dump(data,file)
    start = end
    end += 500
    time.sleep(1)

data = []

for file in files:
    with open(file,'rb') as f:
        info = pkl.load(f)
    tickers = list(info)
    points = ['symbol', 'netProfitMarginTTM',
              'peRatio', 'pegRatio', 'high52', 'currentRatio', 'quickRatio', 'interestCoverage', 'pcfRatio', "divGrowthRate3Year", "returnOnEquity"]
    for ticker in tickers:
        tick = []
        for point in points:
            tick.append(info[ticker]['fundamental'][point])
        data.append(tick)
    os.remove(file)

points = ['symbol', 'Margin', 'PE', 'PEG', 'high52', 'Current Ratio','Quick Ratio','Interest Coverage','Price to Cash Flow','Dividend Growth','ROE']

df_results = pd.DataFrame(data,columns=points)

#df_peg =  df_results[df_results['PEG'] > 1]

df_peg = df_results[(df_results['PEG'] < 1) & (df_results['PEG'] > 0) & (df_results['Margin'] > 20) & (df_results['PE'] > 5) 
                    & (df_results['Price to Cash Flow'] > 0) & (df_results['Interest Coverage'] > 1.5) & (df_results['Current Ratio'] > .8) 
                    & (df_results['Dividend Growth'] >= 0) & (df_results['ROE'] >= 8)]

#print(df_peg)

"""def view(size):
    start = 0 
    stop = size
    while stop < len(df_peg):
        print(df_peg[start:stop])
        start = stop
        stop += size
    print(df_peg[start:stop]) """       

pd.set_option('display.max_rows',1000)

df_peg = df_peg.sort_values(['PEG'])

#Return just the company name and information for what was selected
df_symbols = df_peg['symbol'].tolist()
new =  df['Symbol'].isin(df_symbols)

companies = df[new]

#@print(companies)

filename = f'watch_list.xlsx'

with pd.ExcelWriter(filename, engine="openpyxl", mode='w') as writer:
    companies.to_excel(
        writer, sheet_name='watch_list',index="False")

wb = load_workbook(filename)

ws = wb.worksheets[0]

wb.save(filename)
