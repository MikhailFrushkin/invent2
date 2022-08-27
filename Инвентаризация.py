import csv
import os
import sys
import time

import pandas as pd
from loguru import logger


def file_name() -> tuple:
    """нахождение файлов с 6.1 и результата просчета"""
    file_list = os.listdir()
    file_base = 'Нет подходящих файлов'
    file_check = 'Нет подходящих файлов'
    for item in file_list:
        if item.endswith('.xlsx'):
            if item.startswith('6.1'):
                file_base = item
            elif item != 'Общий итог.xlsx':
                file_check = item

    logger.info('\nФайл из 6.1: {}\nФайл проверки: {}'.format(
        file_base, file_check
    ))
    return file_base, file_check


def read_file(names: tuple):
    """Преобразовавние в csv"""
    try:
        excel_data_df = pd.read_excel('{}'.format(names[0]),
                                      sheet_name='6.1 Складские лоты', skiprows=13, header=1)

        excel_data_df.to_csv('base.csv', encoding='utf-8-sig')

        excel_data_df = pd.read_excel('{}'.format(names[1]),
                                      sheet_name='Sheet1', header=0)

        excel_data_df.to_csv('check.csv', encoding='utf-8-sig')

    except Exception as ex:
        logger.debug('Ошибка открытия файла: {}'.format(ex))
        time.sleep(30)


def comparison():
    """Основная функция"""
    result_list = list()
    art_list = list()
    result = list()
    try:
        print('Считываю файл склада 6.1...')
        with open('base.csv', newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['Код \nноменклатуры'] not in art_list:
                    art_list.append(row['Код \nноменклатуры'])
                    result_list.append([row['Местоположение'],
                                        row['Код \nноменклатуры'],
                                        row['Описание товара'],
                                        int(row['Физические \nзапасы'].replace('.0', '')
                                            if row['Физические \nзапасы'] != ''
                                            else row['Физические \nзапасы'].replace('', '0')),
                                        int(row['Зарезерви\nровано'].replace('.0', '')
                                            if row['Зарезерви\nровано'] != ''
                                            else row['Зарезерви\nровано'].replace('', '0')),
                                        int(row['Доступно'].replace('.0', '')
                                            if row['Доступно'] != ''
                                            else row['Доступно'].replace('', '0')),
                                        0])
    except Exception as ex:
        logger.debug(ex)
        time.sleep(30)

    try:
        print('Считываю файл просчета...')
        with open('check.csv', newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            check_cells = []
            check_art = []
            for row in reader:
                if row['Местоположение'] not in check_cells:
                    check_cells.append(row['Местоположение'])
                if row['Код номенклатуры'] not in check_art:
                    check_art.append(row['Код номенклатуры'])

        print('Сверка, ожидайте...')

        print('Расхождения:')

        for item in check_cells:
            with open('check.csv', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    for i in result_list:
                        if row['Местоположение'] == item:
                            if row['Код номенклатуры'] in art_list and row['Код номенклатуры'] == i[1] and i[0] == item:
                                i[-1] = int(row['Количество факт'])
                                result.append(i)
                                logger.info(i)
                            elif row['Код номенклатуры'] not in art_list \
                                    and row['Код номенклатуры'] not in [q[1] for q in result]:
                                result.append([item, row['Код номенклатуры'], '', 0, 0, 0, int(row['Количество факт'])])
                                logger.info([item, row['Код номенклатуры'], '', 0, 0, 0, int(row['Количество факт'])])
                            elif i[0] == item and i[1] not in check_art and i[1] not in [w[1] for w in result]:
                                result.append(i)
                                logger.info(i)

        for i in sorted(result):
            delta = i[-1] - i[3]
            i.append(delta)

        time.sleep(1)
        print('Запись результатов...')
        write_result(result, 'Общий итог')
        time.sleep(1)
        write_exsel('Общий итог')

    except Exception as ex:
        logger.debug(ex)
        time.sleep(30)
    finally:
        os.remove('check.csv')
        os.remove('base.csv')
        print('Завершено!')
        time.sleep(120)


def align_left(x):
    return ['text-align: left' for x in x]


def write_exsel(name):
    try:
        df = pd.read_csv('{}.csv'.format(name), encoding='utf-8-sig', delimiter=";")
        writer = pd.ExcelWriter('{}.xlsx'.format(name))
        df.style.apply(align_left, axis=0).to_excel(writer, sheet_name='Sheet1', index=False, na_rep='NaN')
        writer.sheets['Sheet1'].set_column(0, 8, 20)
        writer.save()
    except Exception as ex:
        logger.debug(ex)
    finally:
        os.remove('{}.csv'.format(name))


def write_result(result: list, name: str):
    """запись расхождений в csv"""
    try:
        with open('{}.csv'.format(name), 'w', newline='', encoding='utf-8-sig') as file:
            file_writer = csv.writer(file, delimiter=";", lineterminator="\r")
            file_writer.writerow(["Местоположение",
                                  "Артикул",
                                  "Наименование",
                                  "Физ.запас",
                                  "В резерве",
                                  "Доступно",
                                  "Посчитано",
                                  "Разница"])

            for i in result:
                file_writer.writerow(i)
    except Exception as ex:
        logger.debug(ex)
        time.sleep(30)


def main():
    read_file(file_name())
    comparison()


if __name__ == "__main__":
    logger.add(sys.stderr, format="{time} {level} {message}", filter="my_module", level="INFO")
    logger.add("logs/file_X.log", retention="7 days")
    main()
