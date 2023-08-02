from config import *
import os
import json
import csv
import shutil
import requests
from openpyxl import Workbook


def get_json():
    response = requests.get(
        url=f'https://search-maps.yandex.ru/v1/?text={text}&type=biz&lang=ru_RU&apikey={apikey}&results={results}')

    status = True

    if response.status_code == 200:

        try:
            shutil.rmtree(
                'data')  # После первого запуска появляется папка data. Ее нужно очищать перед получением данных!
        except:
            print('data не существует')

        os.makedirs(f'data/{text}')

        data = response.json()

        answer = f'По запросу "{data["properties"]["ResponseMetaData"]["SearchRequest"]["request"]}" найдено {data["properties"]["ResponseMetaData"]["SearchResponse"]["found"]} организаций!'
        print(answer)

        with open(f'data/{text}/data.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
    else:
        print("Ошибка при выполнении запроса:", response.status_code)
        status = False

    return status


def get_data():
    status = get_json()
    if status == True:
        all_data = []

        with open(f'data/{text}/data.json', encoding='utf-8') as file:
            data = json.load(file)

        data = data["features"]
        all_properties = [item['properties'] for item in data if 'properties' in item]

        count = 0
        for organization in all_properties:
            count = count + 1

            with open(f'data/{text}/data{count}.json', 'w', encoding='utf-8') as file:
                json.dump(organization, file, indent=4, ensure_ascii=False)

            organization = organization["CompanyMetaData"]

            try:
                id = organization["id"]
            except:
                id = ' '

            try:
                name = organization["name"]
            except:
                name = ' '

            try:
                address = organization["address"]
            except:
                address = ' '

            try:
                url = organization["url"]
            except:
                url = ' '

            try:
                categories = '/ '.join([item['name'] for item in organization["Categories"] if 'name' in item])
            except:
                categories = ' '
            try:
                phone = organization["Phones"][0]["formatted"]

            except:
                phone = ' '

            try:
                hours = organization["Hours"]["text"]
            except:
                hours = ''

            print(
                f'id: {id} \nname: {name} \naddress: {address} \nurl: {url} \ncategories: {categories} \nphone: {phone} \nhours: {hours} \n ')

            data = [id, name, address, url, categories, phone, hours]
            all_data.append(data)

        return all_data


def get_csv_xlsx():
    all_data = get_data()

    with open(f'data/{text}.csv', 'w', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=",", lineterminator="\r")
        writer.writerow(['Код', 'Название', 'Адрес', 'Сайт', 'Категория', 'Телефон', 'График'])


    for data in all_data:
        with open(f'data/{text}.csv', 'a', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=",", lineterminator="\r")
            writer.writerow(data)

    csv_data = []

    with open(f'data/{text}.csv', "r", encoding="utf-8") as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            csv_data.append(row)

    workbook = Workbook()
    sheet = workbook.active

    for row_index, row in enumerate(csv_data, 1):
        for column_index, value in enumerate(row, 1):
            sheet.cell(row=row_index, column=column_index, value=value)

    workbook.save(f'data/{text}.xlsx')


def main():
    get_json()
    get_data()
    get_csv_xlsx()
    print('Успех!')


if __name__ == '__main__':
    main()
