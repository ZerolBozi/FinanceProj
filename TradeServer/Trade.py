import ccxt

def open_market_order(
        binance: ccxt.binance, 
        symbol: str,
        side: str,
        amount: float,
        stop_loss_price=None,
        take_profit_price=None,
        margin_mode=None,
        leverage=None
    ):
    """
    :param symbol: item name
    :type symbol: string "BTC/USDT"

    :param side: 'buy' or 'sell' (long or short)
    :type side: string

    :param amount: order amount
    :type amount: float

    :param stop_loss_price: set stop loss price
    :type stop_loss_price: float

    :param take_profit_price: set take profit price
    :type take_profit_price: float

    :param margin_mode: 'cross' or 'isolated' (if margin_mode == None then use default margin mode)
    :type margin_mode: string

    :param leverage: set leverage (if leverage == None then use default leverage)
    :type leverage: float

    :rtype: if successed return order info
    """
    if leverage is not None:
        binance.set_leverage(leverage=leverage,symbol=symbol)

    if margin_mode is not None:
        binance.set_margin_mode(marginMode=margin_mode,symbol=symbol)

    ticker = binance.fetch_ticker(symbol)
    current_price = ticker['last']

    amount = round(amount / current_price, 3)

    try:
        order = binance.create_market_order(symbol, side, amount)

        if stop_loss_price:
            stop_loss_order = binance.create_order(
                symbol=symbol,
                type='STOP_MARKET',
                side='sell' if side == 'buy' else 'buy',
                amount=amount,
                params={'stopPrice': stop_loss_price}
            )

        if take_profit_price:
            take_profit_order = binance.create_order(
                symbol=symbol,
                type='TAKE_PROFIT_MARKET',
                side='sell' if side == 'buy' else 'buy',
                amount=amount,
                params={'stopPrice': take_profit_price}
            )
    except ccxt.InsufficientFunds:
        return {
            "msg": "InsufficientFunds",
            'order': '',
            'stop_loss_order': '',
            'take_profit_order': ''
        }
    
    except ccxt.InvalidOrder:
        return {
            "msg": "InvalidOrder",
            'order': '',
            'stop_loss_order': '',
            'take_profit_order': ''
        }
    
    except Exception as e:
        return {
            "msg": 'error' + str(e),
            'order': '',
            'stop_loss_order': '',
            'take_profit_order': ''
        }

    return {
        "msg": "Successed",
        'order': order['id'],
        'stop_loss_order': stop_loss_order['id'] if stop_loss_price else '',
        'take_profit_order': take_profit_order['id'] if take_profit_price else ''
    }
    
def open_limit_order(
        binance: ccxt.binance,
        symbol: str,
        side: str,
        amount: float,
        limit_price: float,
        stop_loss_price=None,
        take_profit_price=None,
        margin_mode=None,
        leverage=None
    ):
    """
    :param symbol: item name
    :type symbol: string

    :param side: 'buy' or 'sell' (long or short)
    :type side: string

    :param amount: order amount
    :type amount: float
    
    :param limit_price: limit price
    :type limit_price: float

    :param stop_loss_price: set stop loss price
    :type stop_loss_price: float

    :param take_profit_price: set take profit price
    :type take_profit_price: float

    :param margin_mode: 'cross' or 'isolated' (if margin_mode == None then use default margin mode)
    :type margin_mode: string

    :param leverage: set leverage (if leverage == None then use default leverage)
    :type leverage: float

    :rtype: if successed return order info
    """
    if leverage is not None:
        binance.set_leverage(leverage=leverage,symbol=symbol)

    if margin_mode is not None:
        binance.set_margin_mode(marginMode=margin_mode,symbol=symbol)

    amount = round(amount / limit_price, 3)
    order = binance.create_limit_order(symbol, side, amount, limit_price)

    try:
        if stop_loss_price:
            stop_loss_order = binance.create_order(
                symbol=symbol,
                type='STOP_MARKET',
                side='sell' if side == 'buy' else 'buy',
                amount=amount,
                params={'stopPrice': stop_loss_price}
            )

        if take_profit_price:
            take_profit_order = binance.create_order(
                symbol=symbol,
                type='TAKE_PROFIT_MARKET',
                side='sell' if side == 'buy' else 'buy',
                amount=amount,
                params={'stopPrice': take_profit_price}
            )
    except ccxt.InsufficientFunds:
        return {
            "msg": "InsufficientFunds",
            'order': '',
            'stop_loss_order': '',
            'take_profit_order': ''
        }
    
    except Exception:
        return {
            "msg": "Error",
            'order': '',
            'stop_loss_order': '',
            'take_profit_order': ''
        }

    return {
        "msg": "Successed",
        'order': order['id'],
        'stop_loss_order': stop_loss_order['id'] if stop_loss_price else '',
        'take_profit_order': take_profit_order['id'] if take_profit_price else ''
    }

def close_market_order(
        binance: ccxt.binance,
        symbol: str,
        side: str,
        amount: float,
    ):
    """
    :param symbol: item name
    :type symbol: string

    :param side: 'buy' or 'sell' (long or short)
    :type side: string

    :param amount: order amount
    :type amount: float

    :rtype: if successed return order info
    """
    ticker = binance.fetch_ticker(symbol)
    current_price = ticker['last']

    amount = round(amount / current_price, 3)

    order = binance.create_market_order(symbol, side, amount)

    post_cancel_order(binance, symbol)

    return {
        "msg": "Successed",
        'order': order['id']
    }

def close_limit_order(
        binance: ccxt.binance,
        symbol: str,
        side: str,
        amount: float,
        limit_price: float
    ):
    """
    :param symbol: item name
    :type symbol: string

    :param side: 'buy' or 'sell' (long or short)
    :type side: string

    :param amount: order amount
    :type amount: float

    :param limit_price: limit price
    :type limit_price: float

    :rtype: if successed return order info
    """
    amount = round(amount / limit_price, 3)
    
    order = binance.create_limit_order(symbol, side, amount, limit_price)

    post_cancel_order(binance, symbol)

    return {
        "msg": "Successed",
        'order': order['id']
    }

def post_cancel_order(
        binance: ccxt.binance,
        symbol: str,
        order_id: str = None
    ):
    """
    :param: symbol: item name
    :type: symbol: string

    :param order_id: the order id
    :type order_id: string

    :rtype: a list of order
    """
    if order_id is None:
        order_list = binance.cancel_all_orders(symbol=symbol)
        return order_list
    order_list = binance.cancel_order(id=order_id,symbol=symbol)
    return order_list

def set_margin_mode(binance: ccxt.binance, marginMode: str, symbol: str):
    """
    :param marginMode: 'cross' or 'isolated'
    :type marginMode: string

    :param symbol: item name
    :type symbol: string
    """
    return binance.set_margin_mode(marginMode=marginMode,symbol=symbol)['msg']

def set_leverage(binance: ccxt.binance, symbol: str, leverage: int):
    """
    :param leverage: leverage
    :type leverage: float

    :param symbol: item name
    :type symbol: string
    """
    return binance.set_leverage(symbol=symbol, leverage=leverage)['leverage']

def get_balance(binance: ccxt.binance) -> dict:
    """
    :rtype: a dict of account balance
    """
    balance = binance.fetch_balance()
    return {
        'total' : balance['total']['USDT'],
        'free': balance['free']['USDT'],
        'used': balance['used']['USDT']
    }

def get_order_status(
        binance: ccxt.binance,
        symbol: str,
        order_id: str
    ):
    """
    0 = NEW
    1 = PARTIALLY_FILLED
    2 = FILLED
    3 = CANCELED
    4 = REJECTED
    5 = EXPIRED
    """
    status = binance.fetch_order(order_id,symbol)['info']['status']
    if status == 'NEW':
        return 0
    if status == 'PARTIALLY_FILLED':
        return 1
    if status == 'FILLED':
        return 2
    if status == 'CANCELED':
        return 3
    if status == 'REJECTED':
        return 4
    if status == 'EXPIRED':
        return 5
    return -1

def get_market_symbols(binance: ccxt.binance):
    markets = binance.load_markets()
    future_markets = {symbol: market for symbol, market in markets.items() if market['type'] == 'swap'}
    return future_markets

def get_market_default_type(binance: ccxt.binance):
    return binance.options['defaultType']

def get_positions(binance: ccxt.binance, symbols: list = None, get_open_orders=False):
    """
    :param symbols: item name
    :type symbols: list
    """
    position_list = binance.fetch_account_positions(symbols=symbols)    
    if get_open_orders:
        return [position for position in position_list if position['contracts'] > 0]
    return position_list

def get_positions_risk(binance: ccxt.binance, symbols: list = None):
    """
    :param symbols: item name
    :type symbols: list

    :rtype: a list of position
    """

    """
    [{
        'info': {
            'symbol': 'BTCUSDT', 
            'positionAmt': '-0.001', 
            'entryPrice': '22345.4', 
            'markPrice': '22338.92789803', 
            'unRealizedProfit': '0.00647210', 
            'liquidationPrice': '221455.40084661', 
            'leverage': '1', 
            'maxNotionalValue': '5.0E8', 
            'marginType': 'cross', 
            'isolatedMargin': '0.00000000', 
            'isAutoAddMargin': 'false', 
            'positionSide': 'BOTH', 
            'notional': '-22.33892789', 
            'isolatedWallet': '0', 
            'updateTime': '1677924853503'
        }, 
        'id': None, 
        'symbol': 'BTC/USDT:USDT', 
        'contracts': 0.001, 
        'contractSize': 1.0, 
        'unrealizedPnl': 0.0064721, 
        'leverage': 1.0, 
        'liquidationPrice': 221455.40084661, 
        'collateral': 199.99582244, 
        'notional': 22.33892789, 
        'markPrice': 22338.92789803, 
        'entryPrice': 22345.4, 
        'timestamp': 1677924853503, 
        'initialMargin': 22.33892789, 
        'initialMarginPercentage': 1.0, 
        'maintenanceMargin': 0.08935571156, 
        'maintenanceMarginPercentage': 0.004, 
        'marginRatio': 0.0004, 
        'datetime': '2023-03-04T10:14:13.503Z', 
        'marginMode': 'cross', 
        'marginType': 'cross', 
        'side': 'short', 
        'hedged': False, 
        'percentage': 0.02
    }] 
    """
    position_list = binance.fetch_positions_risk(symbols=symbols)
    return position_list

def get_max_leverage(binance: ccxt.binance, symbol: str) -> float:
    leverage_info = binance.fetch_leverage_tiers([symbol])
    return leverage_info[symbol.replace("USDT", "/USDT:USDT")][0]['maxLeverage']