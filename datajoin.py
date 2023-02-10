# MIT License
#
# Copyright (c) 2023-01-28 Alexander Zakharov (adzakha2@mts.ru)
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import csv
import fnmatch
import logging
import os
import sys
from argparse import Namespace
from typing import TextIO, Type, Union
from argparse import ArgumentParser

from dateutil import parser as dtparser


HEADER = ['Скорость, об/мин', 'Скорость сдвига, с^-1', 'Вязкость, прямой ход; сПз',
          'Напряжение сдвига, прямой ход; дин/см^2', 'Вязкость, обратный ход; сПз',
          'Напряжение сдвига, обратный ход; дин/см^2']


def sniff_dialect(fp: TextIO) -> Type[csv.Dialect]:
    """
    Определяет диалект CSV

    :param fp: Объект файла
    :return: Диалект CSV
    """
    try:
        dialect = csv.Sniffer().sniff(fp.read(2000), delimiters=',;\t')
        return dialect
    except csv.Error:
        return csv.excel
    finally:
        fp.seek(0)


def transform_value(value: str) -> Union[str, int, float]:
    """
    Преобразует значения в int/float если это возможно

    :param value: Значение ячейки
    :return: Преобразованное значение
    """
    try:
        return int(value)
    except ValueError:
        try:
            return float(value)
        except ValueError:
            return value


def read_data(filename: str) -> dict:
    """
    Считывает данные из файла

    :param filename: Имя файла
    :return: Словарь со значениями вязкости
    """
    with open(filename, encoding='cp1251') as fp:
        data = list(csv.reader(fp, dialect=sniff_dialect(fp)))
        dt = dtparser.parse(f'{data[7][2]} {data[7][3]}')
        return dict(DT=dt, **dict(zip(data[-3], map(lambda v: transform_value(v), data[-1]))))


def collect_files(directory: str) -> list:
    """
    Создает список файлов для чтения

    :param directory: Имя директории с файлами
    :return: Список путей файлов
    """
    output = []
    for dirname, _dirs, files in os.walk(directory):
        for filename in files:
            if fnmatch.fnmatch(filename, '*.csv') and not fnmatch.fnmatch(filename, 'dj_*'):
                output.append(os.path.join(dirname, filename))
    return output


def stringify_values(value: Union[str, int, float]) -> str:
    """
    Возвращает текстовое представление значения в ячейке. Преобразует float в читаемое Excel значение

    :param value: Значение
    :return: Текстовое значение
    """
    if isinstance(value, float):
        return str(value).replace('.', ',')
    return value


def calculate_rates(points: list) -> list:
    """
    Рассчитывает прямой и обратный ход измерений в единый список

    :param points: Точки измерений вязкости
    :return: Список строк измерений
    """
    output = {}
    sorted_points = sorted(points, key=lambda p: p['DT'])
    order = sorted(list({(point['Speed'], point['Shear Rate']) for point in sorted_points}))
    max_point = order[-1]
    for point in sorted_points:
        point_index = (point['Speed'], point['Shear Rate'])
        if point_index == max_point:
            output[point_index] = [
                point['Speed'],
                point['Shear Rate'],
                point['Viscosity'],
                point['Shear Stress'],
                point['Viscosity'],
                point['Shear Stress']
            ]
            continue
        if point_index in output and len(output[point_index]) < 6:
            output[point_index].extend([point['Viscosity'], point['Shear Stress']])
            continue
        if point_index not in output:
            output[point_index] = [
                point['Speed'],
                point['Shear Rate'],
                point['Viscosity'],
                point['Shear Stress']
            ]

    return [output[index] for index in order]


def write_data(filename: str, points: list):
    """
    Записывает измерения в файл CSV

    :param filename: Имя файла
    :param points: Список строк с измерениями
    """
    with open(filename, 'w', newline='') as fp:
        writer = csv.writer(fp, dialect=csv.excel)
        writer.writerow(HEADER)
        for point in points:
            writer.writerow([stringify_values(v) for v in point])


def dump(filename: str, points: list):
    """
    Сохраняет список точек измерений "как есть"

    :param filename: Имя файла
    :param points: Список с точками измерений
    """
    header = list(points[0].keys())
    with open(filename, 'w') as fp:
        writer = csv.writer(fp, dialect=csv.excel)
        writer.writerow(header)
        writer.writerows([[stringify_values(v) for v in point.values()] for point in points])


def main(options: Namespace):
    """Точка вхождения"""
    log = logging.getLogger()
    log.info(f'Обработка файлов в директории "{options.directory}"')
    points = []
    for filename in sorted(collect_files(options.directory)):
        log.info(f'Получение данных из файла {os.path.basename(filename)}')
        points.append(read_data(filename))
    if options.dump:
        log.info(f'Дамп точек данных в файл {options.output}')
        dump(options.dump, points)
    log.info(f'Запись изменений в файл {options.output}')
    write_data(options.output, calculate_rates(points))
    return 0


if __name__ == '__main__':
    parser = ArgumentParser(prog='datajoin', description='Утилита сборки отчета с вискозиметра Brookfield dv2')
    parser.add_argument('directory', help='Директория с данными (по-умолчанию - текущая)',
                        default=os.getcwd(), nargs='?')
    parser.add_argument('-o', '--output', help='Имя файла для вывода', default='dj_joined.csv')
    parser.add_argument('-d', '--dump', help='Имя файла для дампа всех точек данных', default='dj_dump.csv')
    args = parser.parse_args(sys.argv[1:])

    logging.basicConfig(format='%(asctime)s - [%(levelname)s]: %(message)s', level=logging.INFO)
    sys.exit(main(args))
