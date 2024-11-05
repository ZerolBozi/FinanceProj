from datetime import datetime

import ccxt
from robyn import Robyn, Request, jsonify

from Fetch import fetch_klines, fetch_klines_by_n, fetch_funding_rate_now
from Trade import (
    open_market_order, open_limit_order, close_market_order, close_limit_order,
    get_positions, get_balance, get_market_symbols, get_max_leverage, 
    set_margin_mode, set_leverage
)

app = Robyn(__file__)

global binances
binances = dict()

@app.get("/future/positions")
async def positions(request: Request):
    query_data = request.query_params.to_dict()
    uid = query_data['uid'][0]
    symbol = query_data.get('symbol', [None])[0]

    binance = binances[uid]['future']
    positions = get_positions(binance, symbol, True)

    all_positions = []
    for position in positions:
        position_dict = {
            'symbol': position['info']['symbol'],
            'side': position['side'],
            'leverage': position['info']['leverage'],
            'amount': str(position['notional']),
            'price': position['info']['entryPrice'],
            'liquidationPrice': position['liquidationPrice'],
            'marginRatio': str(position['marginRatio']),
            'margin': position['info']['isolatedWallet'],
            'unrealizedPnl': position['info']['unrealizedProfit'],
            'percentage': str(position['percentage']),
        }

        all_positions.append(position_dict)

    if all_positions == []:
        all_positions.append({
            'symbol': '',
            'side': '',
            'leverage': '',
            'amount': '',
            'price': '',
            'liquidationPrice': '',
            'marginRatio': '',
            'margin': '',
            'unrealizedPnl': '',
            'percentage': '',
        })

    return jsonify(all_positions)

@app.post("/future/marketOrder/open")
async def market_order(request: Request):
    query_data = request.json()
    uid = query_data['uid']
    symbol = query_data['symbol']
    side = query_data['side']
    amount = query_data['amount']

    sl_price = query_data.get('SLPrice', None)
    tp_price = query_data.get('TPPrice', None)

    margin_mode = query_data.get('marginMode', None)
    leverage = query_data.get('leverage', None)

    binance = binances[uid]['future']

    amount = float(amount)
    sl_price = float(sl_price) if sl_price != '' else None
    tp_price = float(tp_price) if tp_price != '' else None
    leverage = int(leverage) if leverage != '' else None

    order_info = open_market_order(binance, symbol, side, amount, sl_price,  tp_price, margin_mode, leverage)

    return jsonify(order_info)

@app.post("/future/limitOrder/open")
async def limit_order(request: Request):
    query_data = request.json()
    uid = query_data['uid']
    symbol = query_data['symbol']
    side = query_data['side']
    amount = query_data['amount']
    price = query_data['price']

    sl_price = query_data.get('SLPrice', None)
    tp_price = query_data.get('TPPrice', None)

    margin_mode = query_data.get('marginMode', None)
    leverage = query_data.get('leverage', None)

    binance = binances[uid]['future']

    amount = float(amount)
    price = float(price)
    sl_price = float(sl_price) if sl_price != '' else None
    tp_price = float(tp_price) if tp_price != '' else None
    leverage = int(leverage) if leverage != '' else None

    order_info = open_limit_order(binance, symbol, side, amount, price, sl_price, tp_price, margin_mode, leverage)

    return jsonify(order_info)

@app.post("/future/marketOrder/close")
async def c_market_order(request: Request):
    query_data = request.json()
    uid = query_data['uid']
    symbol = query_data['symbol']
    side = query_data['side']
    amount = query_data['amount']

    binance = binances[uid]['future']

    amount = float(amount)

    order_info = close_market_order(binance, symbol, side, amount)

    return jsonify(order_info)

@app.post("/future/limitOrder/close")
async def c_limit_order(request: Request):
    query_data = request.json()
    uid = query_data['uid']
    symbol = query_data['symbol']
    side = query_data['side']
    amount = query_data['amount']
    price = query_data['price']

    binance = binances[uid]['future']

    amount = float(amount)
    price = float(price)

    order_info = close_limit_order(binance, symbol, side, amount, price)

    return jsonify(order_info)

@app.get("/future/maxLeverage")
async def max_leverage(request: Request):
    query_data = request.query_params.to_dict()

    uid = query_data['uid'][0]
    symbol = query_data['symbol'][0]

    binance = binances[uid]['future']
    max_leverage = get_max_leverage(binance, symbol)

    return jsonify({"maxLeverage": max_leverage})

@app.get("/future/symbols")
async def future_symbols(request: Request):
    query_data = request.query_params.to_dict()

    uid = query_data['uid'][0]

    binance = binances[uid]['future']
    symbols = get_market_symbols(binance)

    return jsonify(symbols)

@app.post("/future/marginMode")
async def margin_mode(request: Request):
    query_data = request.json()

    uid = query_data['uid']
    symbol = query_data['symbol']
    margin_mode = query_data['marginMode']

    margin_mode = 'isolated' if margin_mode == '' else margin_mode

    binance = binances[uid]['future']
    result = set_margin_mode(binance, margin_mode, symbol)

    return jsonify({"msg": result})

@app.post("/future/leverage")
async def margin_mode(request: Request):
    query_data = request.json()

    uid = query_data['uid']
    symbol = query_data['symbol']
    leverage = query_data['leverage']

    leverage = 1 if leverage == '' else leverage

    binance = binances[uid]['future']
    result = set_leverage(binance, symbol, int(leverage))

    return jsonify({"msg": result})

@app.get("/logout")
async def logout(request: Request):
    query_data = request.query_params.to_dict()
    uid = query_data['uid'][0]

    if uid in binances.keys():
        binances.pop(uid)
        return jsonify({
            'status': 'success'
        })
    else:
        return jsonify({
            'status': 'failed'
        })

@app.post("/login")
async def login(request: Request):
    query_data = request.json()
    apikey = query_data.get('apikey', '')
    secret = query_data.get('secret', '')

    balance_dict = {
        'uid': '',
        'total' : '',
        'free': '',
        'used': ''
    }

    global binances

    if apikey == '' or secret == '':
        return jsonify(balance_dict)
    
    binance = ccxt.binance({
        'apiKey': apikey,
        'secret': secret,
        'enableRateLimit': True,
    })

    binance_future = ccxt.binance({
        'apiKey': apikey,
        'secret': secret,
        'enableRateLimit': True,
        'options': {
            'defaultType': 'future'
        }
    })

    try:
        binance.load_markets()

    except Exception as e:
        print(e)
        return jsonify(balance_dict)

    balance = binance.fetch_balance()

    balance_dict = {
        'uid': balance['info'].get('uid', ''),
        'total' : balance['total']['USDT'],
        'free': balance['free']['USDT'],
        'used': balance['used']['USDT']
    }

    binances.setdefault(balance['info']['uid'], {'spot': binance, 'future': binance_future})

    return jsonify(balance_dict)

@app.get("/balance")
async def balance(request: Request):
    query_data = request.query_params.to_dict()

    uid = query_data['uid'][0]
    market_type = query_data.get('market', ['spot'])[0]

    binance = binances[uid][market_type]

    balance_dict = get_balance(binance)

    return jsonify(balance_dict)

@app.get("/assets")
async def assets(request: Request):
    query_data = request.query_params.to_dict()

    uid = query_data['uid'][0]

    binance = binances[uid]['spot']
    balance = binance.fetch_balance()

    filtered_balances = [asset for asset in balance['info']['balances'] if float(asset['free']) > 0 or float(asset['locked']) > 0]
    
    return jsonify(filtered_balances)

@app.get("/fetch/klines")
async def fetch_candles(request: Request):
    query_data = request.query_params.to_dict()
    
    uid = query_data['uid'][0]
    market_type = query_data.get('market', ['spot'])[0]
    symbol = query_data.get('symbol', ['BTC/USDT'])[0]
    interval = query_data.get('timeframe', ['1d'])[0]
    binance = binances[uid][market_type]
    
    start = datetime.now() if 'start' not in query_data else datetime.strptime(query_data['start'][0], '%Y-%m-%d')
    end = datetime.now() if 'end' not in query_data else datetime.strptime(query_data['end'][0], '%Y-%m-%d')

    data = fetch_klines(binance, symbol, interval, start, end, 0.3, need_datetime=True, format_json=True)
    return jsonify(data)

@app.get("/fetchNow/klines")
async def fetch_klines_now(request: Request):
    query_data = request.query_params.to_dict()
    
    uid = query_data['uid'][0]
    market_type = query_data.get('market', ['spot'])[0]
    symbol = query_data.get('symbol', ['BTC/USDT'])[0]
    interval = query_data.get('timeframe', ['1d'])[0]
    binance = binances[uid][market_type]

    data = fetch_klines_by_n(binance, symbol, interval, 1, 0.3, need_datetime=True, format_json=True)
    data = data.replace('[', '').replace(']', '')

    return data

@app.get("/fetchNow/fundingRate")
async def fetch_funding_rate(request: Request):
    query_data = request.query_params.to_dict()
    
    uid = query_data['uid'][0]
    symbol = query_data.get('symbol', ['BTC/USDT'])[0]
    symbol = 'BTC/USDT' if symbol == "" else symbol
    binance = binances[uid]['future']

    funding_rate = fetch_funding_rate_now(binance, symbol)['fundingRate']

    return jsonify({
        'fundingRate': funding_rate
    })

if __name__ == "__main__":
    app.start(port=8080)