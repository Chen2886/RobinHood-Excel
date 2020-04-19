import csv
import shelve
import xlsxwriter

from yahoo_fin import stock_info as si
from RobinLib import Robinhood
from Secret import username, password
from Stock_Orders import StockOrder


def get_symbol_from_instrument_url(rb_client, url, db):
    instrument = {}
    if url in db:
        instrument = db[url]
    else:
        db[url] = fetch_json_by_url(rb_client, url)
        instrument = db[url]
    return instrument['symbol']


def fetch_json_by_url(rb_client, url):
    return rb_client.session.get(url).json()


def order_item_info(order, rb_client, db):
    # side: .side,  price: .average_price, shares: .cumulative_quantity, instrument: .instrument, date : .last_transaction_at
    symbol = get_symbol_from_instrument_url(rb_client, order['instrument'], db)
    return {
        'side': order['side'],
        'price': order['average_price'],
        'shares': order['cumulative_quantity'],
        'symbol': symbol,
        'date': order['last_transaction_at'],
        'state': order['state']
    }


def get_all_history_orders(rb_client):
    orders = []
    past_orders = rb_client.order_history()
    orders.extend(past_orders['results'])
    while past_orders['next']:
        print("{} order fetched".format(len(orders)))
        next_url = past_orders['next']
        past_orders = fetch_json_by_url(rb_client, next_url)
        orders.extend(past_orders['results'])
    print("{} order fetched".format(len(orders)))
    return orders


def get_all_orders(rb):
    past_orders = get_all_history_orders(rb)
    instruments_db = shelve.open('instruments.db')
    orders = [order_item_info(order, rb, instruments_db) for order in past_orders]
    keys = ['side', 'symbol', 'shares', 'price', 'date', 'state']
    with open('orders.csv', 'w') as output_file:
        dict_writer = csv.DictWriter(output_file, keys)
        dict_writer.writeheader()
        dict_writer.writerows(orders)


def auto_adjust_column(worksheet):
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column  # Get the column name
        # Since Openpyxl 2.6, the column name is  ".column_letter" as .column became the column number (1-based)
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width


def write_sheet(workbook, current_sheet, stock_order_list, live_price = 0):
    # for header
    col = 0
    keys = ['Date', 'Symbol', 'Shares', 'Price', 'Total Price']
    for key in keys:
        current_sheet.write(0, col, key, default_format)
        col += 1

    # for all orders
    row = 1
    col = 0
    total_cost = 0
    total_share = 0

    # writing in workbook
    for stock in stock_order_list:
        current_sheet.write(row, col, stock.date, date_format)
        current_sheet.write(row, col + 1, stock.symbol, default_format)
        current_sheet.write(row, col + 2, stock.shares, share_format)
        current_sheet.write(row, col + 3, stock.price, currency_format)
        stock_total = float(stock.shares) * float(stock.price)
        current_sheet.write(row, col + 4, stock_total, currency_format)
        total_cost += stock_total
        total_share += stock.shares
        row += 1

    current_sheet.write(1, 6, "Total Cost", default_format)
    current_sheet.write(2, 6, total_cost, currency_format)
    if "All History" not in current_sheet.get_name():
        current_sheet.write(3, 6, "Total Shares", default_format)
        current_sheet.write(4, 6, total_share, share_format)
        current_sheet.write(3, 7, "Latest Price", default_format)
        current_sheet.write(4, 7, live_price, currency_format)
        current_sheet.write(1, 7, "Equity", default_format)
        current_sheet.write(2, 7, "=G5*H5", currency_format)
        current_sheet.write(1, 8, "Total Profit", default_format)
        current_sheet.write(2, 8, "=H3-G3", currency_format)

    current_sheet.set_column(0, 0, 13)
    current_sheet.set_column(1, 2, 8)
    current_sheet.set_column(3, 20, 13)
    return total_share * live_price


refresh = input("Refresh Data?[y/n]: ")
if "y" in refresh:
    refresh = input("CONFIRM Refresh Data?[y/n]: ")
    if "y" in refresh:
        print("Refreshing Data...")
        robin = Robinhood()
        robin.login(username=username, password=password)
        get_all_orders(robin)

# opening file
file = open("orders.csv", "r")
file.readline()

# list of stock orders
all_orders = []

# list of owned symbols
all_symbols = {}

for order in file:

    # skip cancelled order
    if "cancelled" in order: continue

    # getting rid of newline char at the end
    order = order.replace("\n", "")

    # setting up object
    order_list = order.split(",")
    order_date = order_list[4][0:order_list[4].find("T")].split("-")
    order_obj = StockOrder(True if order_list[0] == "buy" else False, order_list[1], order_list[2], order_list[3],
                           order_date[0], order_date[1], order_date[2])
    all_orders.append(order_obj)

    # adding symbols to list
    if order_list[1] not in all_symbols: all_symbols[order_list[1]] = 0

# New workbook
wb = xlsxwriter.Workbook('Stock History.xlsx')
all_history_sheet = wb.add_worksheet('All History')

# all the format
# format
default_format = wb.add_format({'font_size': 13, 'align': 'center'})
date_format = wb.add_format({'num_format': 'm/d/yy', 'font_size': 13, 'align': 'center'})
share_format = wb.add_format({'num_format': '#,##0;[Red]-#,##0', 'font_size': 13})
currency_format = wb.add_format({'num_format': '_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)',
                                       'font_size': 13})

# write history sheet
write_sheet(wb, all_history_sheet, all_orders)

# loading real time stock data
print("Loading real time stock data...")
print("NOTE: stock data is from the latest non extended hour price")
for symbol in all_symbols:
    all_symbols[symbol] = round(si.get_live_price(symbol), 2)
print(all_symbols)

# entire portfolio
total_equity = 0.0

# writing each sheet
for symbol in all_symbols:
    stock_obj_list = [stock_obj for stock_obj in all_orders if symbol in stock_obj.symbol]
    sheet = wb.add_worksheet(symbol)
    total_equity += write_sheet(wb, sheet, stock_obj_list, all_symbols[symbol])

all_history_sheet.write(1, 7, "Total Equity", default_format)
all_history_sheet.write(2, 7, round(total_equity, 2), currency_format)
all_history_sheet.write(1, 8, "Total Profit", default_format)
all_history_sheet.write(2, 8, "=H3-G3", currency_format)
wb.close()
