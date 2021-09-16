import collections
import datetime
import os
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    load_dotenv()

    current_year = datetime.datetime.now()
    age = current_year.year - 1920
    years = get_plural_years(age)

    file = os.getenv("FILE")
    excel_data_df = pandas.read_excel(
        file,
        sheet_name='Лист1',
        na_values='nan',
        keep_default_na=False,
    )

    wines = excel_data_df.to_dict(orient='records')
    wines_by_category = collections.defaultdict(list)

    for wine in wines:
        wines_by_category[wine['Категория']].append(wine)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        age=f'Уже {age} {years} с вами',
        wines=wines_by_category,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


def get_plural_years(n):
    years = ['год', 'года', 'лет']
    if all((n % 10 == 1, n % 100 != 11)):
        return years[0]
    elif all((2 <= n % 10 <= 4, any((n % 100 < 10, n % 100 >= 20)))):
        return years[1]
    return years[2]



