from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict
from collections import OrderedDict
import os
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import pprint


def group_wines_from_excel(file_path):
    excel_data_df = pandas.read_excel(file_path)
    excel_data_df = excel_data_df.where(pandas.notnull(excel_data_df), None)
    data_dict = excel_data_df.to_dict(orient='records')

    category_wines = defaultdict(list)
    for wine in data_dict:
        wine['is_best_offer'] = 'Выгодное предложение' in str(wine.values())
        category_wines[wine['Категория']].append(wine)

    sorted_categories = sorted(category_wines.keys())
    ordered_wines = OrderedDict()
    for category in sorted_categories:
        ordered_wines[category] = category_wines[category]

    return dict(ordered_wines)


def main():
    load_dotenv()
    file_path = os.getenv('WINE3', default=None)
    category_wines = group_wines_from_excel(file_path)
    
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('wine_shop.html')

    birth_year = 1920
    current_year = datetime.datetime.now().year
    years = current_year - birth_year
    if years % 10 == 1 and years % 100 != 11:
        year_word = "год"
    elif years % 10 in [2, 3, 4] and not (years % 100 in [12, 13, 14]):
        year_word = "года"
    else:
        year_word = "лет"

    rendered_page = template.render(
        years=str(years),
        year_word=year_word,
        data_dict=category_wines,
    )

    with open('output.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
