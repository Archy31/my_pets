import json
import re

import requests
from bs4 import BeautifulSoup
import lxml

from openpyxl import load_workbook

from config import base_url


def get_count_pages():
    respo = requests.get(base_url).text
    soup = BeautifulSoup(respo, 'lxml')
    pages = soup.find(class_='pagination').find_all('li')

    return int(pages.pop().text)


def add_to_excel():
    pn = load_workbook('results/parsed_dates.xlsx')
    try:
        sheet = pn.active

        sheet['A1'] = 'НАЗВАНИЕ'
        sheet['B1'] = 'ЦЕНА'
        sheet['C1'] = 'ОПИСАНИЕ'
        sheet['D1'] = 'ССЫЛКА'

        with open('results/dates.json') as f:
            dates = json.load(f)

        count = 2
        for name, titles in dates.items():
            sheet[f'A{count}'] = name

            sheet[f'B{count}'] = titles['price']
            sheet[f'C{count}'] = titles['avail']
            sheet[f'D{count}'] = titles['link']

            count += 1

    except Exception as e:
        return e

    finally:
        pn.save('results/parsed_dates.xlsx')
        pn.close()


all_dates = {}


def get_shines(url):
    try:
        respo = requests.get(url)
        soup = BeautifulSoup(respo.text, 'lxml')
        prod_s = soup.find(class_='columns-container').find_all(class_='product')
        # prod_s = soup.find(class_='products-list').find_all(class_='product ')
        for prod1 in prod_s:
            try:
                prod = prod1.find(class_='product-bottom')
                name_title = prod.find(class_='name').find('a')
                price_prod = prod.find('span', class_='price').text.strip()
                link_prod = prod.find(class_='name').find('a').get('href')

                avail_prod = prod1.find(class_='availability-container').text.strip().replace(' ', '').replace('\n', ' ')

            except Exception as e:
                return e

            na = re.split(r"[><]+", str(name_title))[2].strip()
            me = re.split(r"[><]+", str(name_title))[4].strip()

            name = re.sub('\xa0', ' ', f'{na} ({me})')
            price = re.sub('\xa0', '', price_prod)
            avail = re.sub('\xa0', '', avail_prod)
            link = re.sub('\xa0', '', link_prod)

            shine_dates = {
                    'price': f'{price}',
                    'avail': f'{avail}',
                    'link': f'{base_url + link}'
                }

            all_dates[name] = shine_dates

    except Exception as e:
        return e


def main():
    count_pages = get_count_pages()

    for i in range(1, count_pages + 1):
        url = base_url + f'?page={i}'
        get_shines(url=url)

    with open('results/dates.json', 'w', encoding='utf-8') as f:
        json.dump(all_dates, f, ensure_ascii=False, indent=4)

    add_to_excel()


if __name__ == '__main__':
    main()
