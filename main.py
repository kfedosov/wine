import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def year_format(years_number):
    last_two_digits = abs(years_number) % 100
    last_digit = abs(years_number) % 10
    if 5 <= years_number <= 20:
        years_string = "лет"
    else:
        if last_digit == 1 and last_two_digits != 11:
            years_string = "год"
        elif 2 <= last_two_digits <= 4:
            years_string = "года"
        else:
            years_string = "лет"
    return years_string


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

now = datetime.datetime.now()
years_passed = now.year - 1920

excel_data_df = pandas.read_excel('wine3.xlsx', sheet_name='Лист1')
# Заменяем nan на None
excel_data_df = excel_data_df.where((pandas.notnull(excel_data_df)), None)
excel_data_dict = defaultdict(list)

for row in excel_data_df.to_dict(orient='records'):
    excel_data_dict[row['Категория']].append(row)

rendered_page = template.render(
    number_of_years=years_passed,
    years_string='лет',
    excel_data_dict=excel_data_dict,
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()