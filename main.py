import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

now = datetime.datetime.now()
age = now.year - 1920


def plural_years(n):
    years = ['год', 'года', 'лет']
    if all((n % 10 == 1, n % 100 != 11)):
        return years[0]
    elif all((2 <= n % 10 <= 4, any((n % 100 < 10, n % 100 >= 20)))):
        return years[1]
    return years[2]
years = plural_years(age)


excel_data_df = pandas.read_excel(
    'wine3.xlsx',
    sheet_name='Лист1',
    na_values='nan',
    keep_default_na=False,
    )


wines = excel_data_df.to_dict(orient='records')
wines_by_category = collections.defaultdict(list)
for wine in wines:
    for k, v in wine.items():
        if k == 'Категория':
            wines_by_category[v].append(wine)
categories = [k for k in wines_by_category.keys()]
categories = sorted(categories)


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

rendered_page = template.render(
    age=f'Уже {age} {years} с вами',
    wines=wines_by_category,
    categories=categories,
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
