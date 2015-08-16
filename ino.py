from bs4 import BeautifulSoup
from collections import namedtuple
from operator import attrgetter
from pprint import pprint
import urllib2 as http
import xlsxwriter

url = "http://quotes.ino.com/exchanges/contracts.html?r=NYMEX_NG"

def pct_to_float(s):
    if s[-1] == '%':
        return float(s[:-1])
    else:
        return float(s)

def parse_table_row(row):
    Row = namedtuple('Row', ['Market', 'Contract', 'Open', 'High', 'Low', 'Last', 'Change', 'Pct'])
    return Row(Market=row[0],
               Contract=row[1],
               Open=parse_to_float(row[2]),
               High=parse_to_float(row[3]),
               Low=parse_to_float(row[4]),
               Last=parse_to_float(row[5]),
               Change=parse_to_float(row[6]),
               Pct=pct_to_float(row[7]))

def parse_to_float(s):
    try:
        return float(s)
    except:
        return float('nan')

def get_url_data(url):
    return http.urlopen(url)

def get_table_header(page):
    soup = BeautifulSoup(page)
    table = soup.find(name = 'table')
    table_header = table.findChild(text='Market').parent.parent
    return table_header

def parse_table(header):
    rows = []
    try: 
        for row_data in header.next_siblings:
            data_strings = [x.text for x in row_data.children]
            row = parse_table_row(data_strings)
            rows.append(row)
    except AttributeError:
        print "Hit end of table data"
    finally:
        pprint(rows)
        return rows

def get_trimmed_contract_data(rows):
     contract_data = data_column(rows, 'Contract')
     return contract_data[:148]

def get_trimmed_price_data(rows):
    price_data = data_column(rows, 'Last')
    return price_data[:148]

def data_column(rows, attr):
    f = attrgetter(attr)
    return [f(x) for x in rows]

def get_trimmed_data(rows, attr):
    f = attrgetter(attr)
    data = [f(x) for x in rows]
    return data[:148]

def two_column_line_chart(column_A, column_B, chart_name='chart_line.xlsx'):
    column_A_start_cell, column_B_start_cell = 'A1', 'B1'
    workbook = xlsxwriter.Workbook(chart_name)
    worksheet = workbook.add_worksheet()
    worksheet.write_column(column_A_start_cell, column_A)
    worksheet.write_column(column_B_start_cell, column_B)
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'categories': '=Sheet1!' + column_A_start_cell + ':A' + str(len(column_A)),
        'values': '=Sheet1!B' + column_B_start_cell + ':B' + str(len(column_B))})
    worksheet.insert_chart('E2', chart)
    workbook.close()

def parse(html):
    header = get_table_header(html)
    rows = parse_table(header)
    contract_data = get_trimmed_data(rows, 'Contract')
    price_data = get_trimmed_data(rows, 'Last')
    return contract_data, price_data

def main():
    html = get_url_data(url)
    contract_data, price_data = parse(html)
    two_column_line_chart(contract_data, price_data)

if __name__ == '__main__':
    main()

