import pandas as pd
import numpy as np
import os
import investpy as inv
import datetime as dt
import pandas_datareader as pdr

ETF_NAME_TO_ETF_ACRONYM = {
'SPY' : 'SPDR S&P 500',
'XLB' : 'Materials Select Sector SPDR',
'XLE' : 'Energy Select Sector SPDR',
'XLF' : 'Financial Select Sector SPDR',
'XLI' : 'Industrial Select Sector SPDR',
'XLK' : 'Technology Select Sector SPDR',
'XLP' : 'Consumer Staples Select Sector SPDR',
'XLU' : 'Utilities Select Sector SPDR',
'XLV' : 'Health Care Select Sector SPDR',
'XLY' : 'Consumer Discretionary Select Sector SPDR',
'XTN' : 'SPDR S&P Transportation',
'EWJ' : 'iShares MSCI Japan',
'EWG' : 'iShares MSCI Germany',
'EEM' : 'iShares MSCI Emerging Markets',
'TLT' : 'iShares 20+ Year Treasury Bond',
'GLD' : 'SPDR Gold Shares'
}

def price_to_returns(df_price):
    return df_price.pct_change().dropna()

date_I = '20/03/2022'
date_II = '03/04/2022'
def etf_historical_returns(etf, date_I, date_II):
    df = inv.get_etf_historical_data(ETF_NAME_TO_ETF_ACRONYM['SPY'],
                                    country = 'united states',
                                    from_date = date_I,
                                    to_date = date_II)['Close']
    returns = pd.DataFrame(price_to_returns(df))
    return returns

    
# historical_prices = pd.concat(df_list, axis = 1)    
def stocks_historical_returns(stock, date_I, date_II):
    historical_prices = pdr.get_data_yahoo(stock,
                                          start = date_I,
                                          end = date_II)['Close']
    returns = pd.DataFrame(price_to_returns(historical_prices))
    return returns
    
def clean_dates(dates):
    dates = dates.split(' - ')
    for num in range(0,len(dates)):
        dates[num] = f'{dates[num]}/2022'
    return dates

def datetime_fmt(date):
    return dt.datetime.strptime(date, '%d/%m/%Y')

def reading_excel_data(excel):
    df = pd.read_excel(excel)
    assets = list(df.iloc[14].dropna())[1:] + list(df.iloc[19].dropna())[1:] 
    weights = list(df.iloc[15].dropna())[1:] + list(df.iloc[20].dropna())[1:]
    group_name = list(df.iloc[8].dropna())[0]
    dates = clean_dates(list(df.iloc[11].dropna())[0])
    return [assets, weights, group_name, dates]

list_archive = []
for _, _, files in os.walk("."):
    for filename in files:
        if filename.endswith('.xlsx'):
            list_archive.append(filename)

def portfolio_stats(weights, df):
    weights = np.array(weights)
    dfr = np.array(((1 + df).cumprod() - 1).iloc[-1])
    annual_return = ((1 + np.dot(weights.T, dfr)) ** (252/5)) - 1
    annual_volatility = np.sqrt(weights.T @ df.cov() * 252 @ weights)
    risk_return = annual_return / annual_volatility
    return {'return': annual_return, 
            'risk': annual_volatility, 
            'risk_return': risk_return}
    
df = pd.DataFrame()
for filename in list_archive:
    group_data = reading_excel_data(filename)
    lst_returns = []
    group_data[3] = ['28/03/2022', '03/04/2022']
    for asset in group_data[0]:
        # is a ETF, investing has to pass as string format
        if len(asset) == 3:
            lst_returns.append(etf_historical_returns(asset, group_data[3][0], group_data[3][1]))
        # is a brStock, yahoo finance has to pass datetime
        else:
            lst_returns.append(stocks_historical_returns(asset, datetime_fmt(group_data[3][0]), datetime_fmt(group_data[3][1])))
    historical_returns = pd.concat(lst_returns, axis = 1)
    historical_returns.columns = group_data[0]
    results = portfolio_stats(group_data[1], historical_returns)['risk_return']
    df = df.append({'Grupo' : group_data[2], 'Risco-Retorno' : results}, ignore_index=True)
df.to_excel('Resultado.xlsx')
            