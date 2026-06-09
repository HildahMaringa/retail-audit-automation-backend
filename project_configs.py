"""
╔══════════════════════════════════════════════════════════════════╗
║  PROJECT CONFIGS — Unified Data Query Engine                     ║
║                                                                  ║
║  One config block per project. The engine reads this file and   ║
║  adapts all logic automatically. To add a new project, copy     ║
║  the closest existing block and update the values.              ║
╚══════════════════════════════════════════════════════════════════╝

price_low / price_high:
    ALL projects now use 0.25 / 2.0 — standardised for consistency.
    0.25 means: current price must be ≥ 25% of historical bound to pass.
"""

PROJECT_CONFIGS = {

    'NG-MRA': {
        'outlet_id':    'outletid',
        'sku_id':       'wh_skuid',
        'channel':      'TradeChannel',
        'price_cols':       ['Capture Price', 'Capture Price Excl Container'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Total Purchase',
        'units_col':                 None,
        'use_units_in_price_comb':   False,
        'extra_checks':          [],
    },

    'Nigeria-Ville': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'SKU_ID',
        'channel':      'Trade Channel',
        'price_cols':       ['Selling Price'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                'Units',
        'use_units_in_price_comb':  True,
        'extra_checks':          [],
    },

    'Pewa': {
        'outlet_id':    'OUTNUMBER',
        'sku_id':       'Prodcode',
        'channel':      'Channel Type',
        'price_cols':       ['Selling Price per Sku'],
        'buying_price_col': 'Buying Price per Sku',
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'TOTALSTOCK',
        'purchase_col': 'PURCHASES',
        'units_col':                None,
        'use_units_in_price_comb':  False,
        'extra_checks':          ['negative_profit', 'no_profit'],
    },

    'Usafi-Uganda': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'wh_skuid',
        'channel':      'Trade Channel',
        'price_cols':       ['Selling Price'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                None,
        'use_units_in_price_comb':  False,
        'extra_checks':          [],
    },

    'Kenya-MRA': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'wh_skuid',
        'channel':      'Trade Channel',
        'price_cols':       ['Selling Price'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                'Units',
        'use_units_in_price_comb':  True,
        'extra_checks':          [],
    },

    'TZ-MRA': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'wh_skuid',
        'channel':      'Trade Channel',
        'price_cols':       ['Selling Price'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                'Units',
        'use_units_in_price_comb':  True,
        'extra_checks':          [],
    },

    'KO-Tanzania': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'wh_skuid',
        'channel':      'Trade Channel',
        'price_cols':       ['Capture Price', 'Capture Price Excl Container'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                None,
        'use_units_in_price_comb':  False,
        'extra_checks':          [],
    },

    'KO-Uganda': {
        'outlet_id':    'Outlet ID',
        'sku_id':       'wh_skuid',
        'channel':      'Trade Channel',
        'price_cols':       ['Capture Price', 'Capture Price Excl Container'],
        'buying_price_col': None,
        'skip_price_val_1': True,
        'price_low':        0.25,
        'price_high':       2.0,
        'stock_col':    'Total Stock',
        'purchase_col': 'Purchases',
        'units_col':                None,
        'use_units_in_price_comb':  False,
        'extra_checks':          [],
    },
}

STOCK_THRESHOLDS = {
    1: (7.0, 1/7),
    2: (5.0, 1/5),
    3: (4.0, 1/4),
    4: (3.0, 1/3),
}

def get_stock_thresholds(mean_val: float):
    if mean_val == 0 or not isinstance(mean_val, (int, float)):
        return 3.0, 1/3
    digits = len(str(int(abs(mean_val)))) if int(abs(mean_val)) > 0 else 1
    return STOCK_THRESHOLDS.get(min(digits, 4), (3.0, 1/3))
