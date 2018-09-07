text = 'kite_spends.txt'

with open(text) as fp:
    values = fp.readlines()

new = [values[idx:idx + 3] for idx in range(0, len(values), 3)]


profit_loss = 0
cur_val = 0
for stock in new:
    stock = [_.replace('\t', ' ').strip() for _ in stock if _.strip()]
    pl = stock[2].split(' ')[0]
    profit_loss += float(pl)
    cur_val += float(stock[1].replace(',', ''))
    # print(stock, pl)

print('Total Stocks =', len(new))
print('Current Stock Value = ', cur_val)
print('Investment Stock Value = ', cur_val - profit_loss)
print('Profit/Loss = ', profit_loss)
