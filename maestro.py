import datetime
from collections import defaultdict

from openpyxl import load_workbook

start = datetime.datetime.now()

COL_PRICE = 18
COL_INVOICE = 26

references = defaultdict(list)
prices = []

input_price_sum = input('Please enter the sum of prices: ')

try:
    price_sum = float(input_price_sum)
except ValueError:
    print('The sum of prices must be a number.')

wb = load_workbook('test.xlsx')
ws = wb.get_active_sheet()

for row in ws.iter_rows(row_offset=1):

    price = row[COL_PRICE].value

    if price is not None and price > 0 and price <= price_sum:

        invoice = row[COL_INVOICE].value
        prices.append(price)
        references[price].append(invoice)


def subsets_with_sum(lst, target):

    x = 1

    def _a(idx, l, r, t):

        if t == sum(l): r.append(l)
        elif t < sum(l): return
        for u in range(idx, len(lst)):
            _a(u + x, l + [lst[u]], r, t)
        return r

    return _a(0, [], [], target)

prices.sort()
print(len(prices))
print(prices)

price_subsets = set([tuple(s) for s in subsets_with_sum(prices, price_sum)])
invoice_subsets = [[references[p] for p in ps] for ps in price_subsets]

print(price_subsets)
print(invoice_subsets)

end = datetime.datetime.now()
print('Time lapsed: ' + str(end - start))
