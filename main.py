import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

WINEMAKING_SINCE_YEAR = 1920


def get_years_string(years_number):
    last_two_digits = abs(years_number) % 100
    last_digit = abs(years_number) % 10
    if 5 <= last_two_digits <= 20:
        return "лет"
    elif last_digit == 1:
        return "год"
    elif 2 <= last_digit <= 4:
        return "года"
    else:
        return "лет"


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    now = datetime.datetime.now()
    years_passed = now.year - WINEMAKING_SINCE_YEAR

    products_data = pandas.read_excel('products.xlsx', sheet_name='Лист1')
    products_data = products_data.where((pandas.notnull(products_data)), None)
    excel_data = defaultdict(list)

    for row in products_data.to_dict(orient='records'):
        excel_data[row['Категория']].append(row)

    rendered_page = template.render(
        number_of_years=years_passed,
        years_string=get_years_string(years_passed),
        excel_data=excel_data,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
