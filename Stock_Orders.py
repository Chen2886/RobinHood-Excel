class StockOrder:
    symbol = ""
    shares = 0.0
    price = 0.0
    year = 0
    month = 0
    day = 0
    date = ""

    def __init__(self, buy, symbol, shares, price, year, month, day):
        self.symbol = symbol
        self.shares = float(shares)
        self.price = float(price)
        self.year = int(year)
        self.month = int(month)
        self.day = int(day)
        self.date = month + "/" + day + "/" + year
        if not buy: self.shares *= -1

    def to_string(self):
        return self.symbol + "|" + str(self.shares) + "|" + str(self.price) + "|" + self.date + "|"



