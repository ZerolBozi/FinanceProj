from time import time, sleep
from datetime import datetime

import ccxt
import pandas as pd

def fetch_funding_rate_now(binance: ccxt.binance, symbol: str):
    return binance.fetch_funding_rate(symbol)

def fetch_klines(
        binance: ccxt.binance, 
        symbol: str, 
        interval: str, 
        start: datetime, 
        end: datetime, 
        delay: float = 0.5, 
        need_datetime: bool = False,
        format_json: bool = False
    ):
    """
    fetch klines from binance

    :param symbol: 商品名稱 (e.g. 'BTC/USDT')
    :type symbol: str

    :param interval: 時間間隔 (e.g. '1m', '5m', '15m', '30m', '1h', '4h', '1d', '1w')
    :type interval: str

    :param start: 開始時間
    :type start: datetime

    :param end: 結束時間
    :type end: datetime

    :param delay: 每次 request 之間的間隔時間
    :type delay: float
    :default delay: 0.5

    :need_datetime: 是否需要 datetime 欄位
    :type need_datetime: bool
    :default need_datetime: False

    :format_json: 是否需要格式化成json
    :type format_json: bool
    :default format_json: False

    :return: klines
    :rtype: pd.DataFrame
    """
    df = pd.DataFrame()
    start_ts = int(start.timestamp() * 1000)
    end_ts = int(end.timestamp() * 1000)
    timeframe_dict = {
        '1m': 60000,
        '5m': 300000,
        '15m': 900000,
        '30m': 1800000,
        '1h': 3600000,
        '4h': 14400000,
        '1d': 86400000,
        '1w': 604800000
    }

    while start_ts <= end_ts:
        data = binance.fetch_ohlcv(symbol, timeframe=interval, since=start_ts, limit=1000)
        
        if len(data) < 0:
            break

        tmp_df = pd.DataFrame(data,columns=['unix','open','high','low','close','volume'])   
        if need_datetime:
            tmp_df.loc[:,'datetime'] = pd.to_datetime(tmp_df.loc[:,'unix'],unit='ms') + pd.Timedelta(hours=8)
            tmp_df = tmp_df.reindex(columns = ['datetime','unix','open','high','low','close', 'volume'])

        df = pd.concat([df, tmp_df], ignore_index=True)
        interval_ms = timeframe_dict.get(interval)
        start_ts = int(df.iloc[-1]['unix']) + interval_ms
        
        sleep(delay)

    df_filtered = df[df['unix'] <= end_ts]

    if format_json:
        return df_filtered.to_json(orient='records', date_format='iso')

    return df_filtered

def fetch_klines_by_n(
        binance: ccxt.binance, 
        symbol: str, 
        interval: str, 
        n: int, 
        delay: float = 0.5,
        need_datetime: bool = False,
        format_json: bool = False
    ):
    now = time()
    timeframe_in_seconds = binance.parse_timeframe(interval)
    start_time_seconds = int(now - n * timeframe_in_seconds)
    start_datetime = datetime.fromtimestamp(start_time_seconds)
    return fetch_klines(binance, symbol, interval, start_datetime, datetime.now(), delay, need_datetime, format_json)