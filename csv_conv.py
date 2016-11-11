#coding=utf-8

import csv
import requests
#from Tkinter import Tk
#from tkFileDialog import askopenfilename, asksaveasfilename
import re
import sys
import logging
from logging import handlers
import socket
import traceback

reload(sys)
sys.setdefaultencoding('utf8')

HOST = ('mail.detmir.ru', 25)
FROM = '"Ошибочка вышла :(" <mlavrikova@detmir.ru>'
TO = 'mlavrikova@detmir.ru', 'ikalinin@detmir.ru'
SUBJECT = 'Ошибка конвертации ERP - ПРОМО'

logging.basicConfig(level=logging.WARNING,
                    format='%(asctime)s.%(msecs)d %(levelname)s %(module)s - %(funcName)s: %(message)s',
                    datefmt="%Y-%m-%d %H:%M:%S",
                    filename=u'converterlog.log')



handler = handlers.SMTPHandler(HOST, FROM, TO, SUBJECT)
email_logger = logging.getLogger('mail.detmir.ru')
email_logger.addHandler(handler)
email_logger.setLevel = logging.CRITICAL


def load_category_csv():
    result = dict()
    with open('category.csv', 'rb') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        for id_erp,id_promo in reader:
            result[id_erp] = id_promo
    return result


category_dict = load_category_csv()


def load_brands():
    response = requests.get('http://www.detmir.ru/api/rest/dictionaries')
    data = response.json()
    result = dict()
    for brand in data['brands']:
        result[brand['title'].replace(' ', '').lower()] = brand['id']
    return result


brands_dict = load_brands()

#Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
#filename = askopenfilename(filetypes=(('Excel files','.xlsx'),('All types', '*.*'))) # show an "Open" dialog box and return the path to the selected file
#print(filename)

regex = re.compile('\(.*?\)')

def floatx(value):
  return float(value or 0.0)

def strx(value):
  return str(value or '')

def get_output_row(irow, is_sup):
    orow = dict()

    # 2-уникальный ID ДМ
    val = unicode(irow['iddm'])
    if is_sup:
        val += u'sup'
    orow['iddm'] = val

    # 3-Артикул
    orow['art'] = irow['art']

    # 5-Наименование
    if irow['name']:
        vals = irow['name'].split(u':')
        if is_sup:
            orow['namepr'] = vals[0].capitalize()
            orow['model'] = vals[1]
        elif len(vals) >= 3:
            orow['namepr'] = vals[0].capitalize()
            orow['model'] = vals[2]
        else:
            orow['namepr'] = strx(irow['name'])

    # 4-Категория
    if irow['groupproduct']:
        id_erp = re.sub(regex, '', irow['groupproduct'])
        if id_erp in category_dict:
            orow['catid'] = category_dict[id_erp]
        else:
            orow['catid'] = irow['groupproduct']
            logging.critical('Не найдено совпадения категории из ERP -' + irow['groupproduct'])
            email_logger.critical('Не найдено совпадения категории из ERP -' + irow['groupproduct'])


    # 6-Бренд
    if irow['brandname']:
        brand = re.sub(regex, '', irow['brandname'])
        if brand and brand.replace(' ', '').lower() in brands_dict:
            orow['brandpr']= brands_dict[brand.replace(' ', '').lower()]
        else:
            orow['brandpr'] = irow['brandname']
            logging.critical('Не найден брэнд -' + irow['brandname'])
            email_logger.critical('Не найден брэнд -' + irow['brandname'])

    # 10-пол
    if irow['sex']:
        if irow['sex'].lower() == 'f':
            orow['sexpr'] = 'ж'
        elif irow['sex'].lower() == 'm':
            orow['sexpr'] = 'м'
        elif irow['sex'].lower() == 'u':
            orow['sexpr'] = ' '
        else:
            orow['sexpr'] = irow['sex']
            logging.warn('Не найдено значение пол -' + irow['sex'])

    # 16-состав товара
    orow['structpr'] = irow['productstr']

    # 18-Возраст от
    if irow['agefromid']:
        if irow['agefromid'].lower() == '0m+':
            orow['agefrompr'] = '0'
        elif irow['agefromid'].lower() == '3m+':
            orow['agefrompr'] = '3'
        elif irow['agefromid'].lower() == '6m+':
            orow['agefrompr'] = '6'
        elif irow['agefromid'].lower() == '9m+':
            orow['agefrompr'] = '9'
        elif irow['agefromid'].lower() == '1,5y+':
            orow['agefrompr'] = '18'
        elif irow['agefromid'].lower() == '10y+':
            orow['agefrompr'] = '120'
        elif irow['agefromid'].lower() == '11y+':
            orow['agefrompr'] = '132'
        elif irow['agefromid'].lower() == '12y+':
            orow['agefrompr'] = '144'
        elif irow['agefromid'].lower() == '1y+':
            orow['agefrompr'] = '12'
        elif irow['agefromid'].lower() == '2y+':
            orow['agefrompr'] = '24'
        elif irow['agefromid'].lower() == '3y+':
            orow['agefrompr'] = '36'
        elif irow['agefromid'].lower() == '5y+':
            orow['agefrompr'] = '60'
        elif irow['agefromid'].lower() == '6y+':
            orow['agefrompr'] = '72'
        elif irow['agefromid'].lower() == '7y+':
            orow['agefrompr'] = '84'
        elif irow['agefromid'].lower() == '8y+':
            orow['agefrompr'] = '96'
        elif irow['agefromid'].lower() == '9y+':
            orow['agefrompr'] = '108'
        elif irow['agefromid'].lower() == 'adult':
            orow['agefrompr'] = ' '
        else:
            orow['agefrompr'] = irow['agefromid']
            logging.warn('Не найден возраст от -' + irow['agefromid'])

    # 19-Возраст до
    if irow['agetoid']:
        if irow['agetoid'].lower() == '3m-':
            orow['agetopr'] = '3'
        elif irow['agetoid'].lower() == '6m-':
            orow['agetopr'] = '6'
        elif irow['agetoid'].lower() == '9m-':
            orow['agetopr'] = '9'
        elif irow['agetoid'].lower() == '1,5y-':
            orow['agetopr'] = '18'
        elif irow['agetoid'].lower() == '10y-':
            orow['agetopr'] = '120'
        elif irow['agetoid'].lower() == '11y-':
            orow['agetopr'] = '132'
        elif irow['agetoid'].lower() == '12y-':
            orow['agetopr'] = '144'
        elif irow['agetoid'].lower() == '1y-':
            orow['agetopr'] = '12'
        elif irow['agetoid'].lower() == '2y-':
            orow['agetopr'] = '24'
        elif irow['agetoid'].lower() == '3y-':
            orow['agetopr'] = '36'
        elif irow['agetoid'].lower() == '4y-':
            orow['agetopr'] = '48'
        elif irow['agetoid'].lower() == '5y-':
            orow['agetopr'] = '60'
        elif irow['agetoid'].lower() == '6y-':
            orow['agetopr'] = '72'
        elif irow['agetoid'].lower() == '7y-':
            orow['agetopr'] = '84'
        elif irow['agetoid'].lower() == '8y-':
            orow['agetopr'] = '96'
        elif irow['agetoid'].lower() == '9y-':
            orow['agetopr'] = '108'
        else:
            orow['agetopr'] = irow['agetoid']
            logging.warn('Не найден возраст до -' + irow['agetoid'])

    #24-Коллекция
    if irow['season']:
        if irow['season'].lower() == 'aw':
            orow['seasonpr'] = 'Осень-зима ' + strx(irow['seasonyear'])
        elif irow['season'].lower() == 'ss':
            orow['seasonpr'] = 'Весна-лето ' + strx(irow['seasonyear'])
        elif irow['season'].lower() == 'sch':
            orow['seasonpr'] = 'Школа ' + strx(irow['seasonyear'])
        elif irow['season'].lower() == 'reg':
            orow['seasonpr'] = strx(irow['seasonyear'])
        else:
            orow['seasonpr'] = strx(irow['season']) + strx(irow['seasonyear'])
            logging.warn('Не найдено совпадение сезона -' + irow['season'])

    #17 - Страна происхождения
    orow['countrypr'] = irow['countryname']

    #12 - Ширина
    orow['wightpr'] = floatx(irow['wight']) * 100.0

    #14 - Высота
    orow['heightpr'] = floatx(irow['height']) * 100.0

    #13 - Длина
    orow['lenghtpr'] = floatx(irow['lenght']) * 100.0

    #15 - Вес
    orow['weightpr'] = floatx(irow['weight'])  * 1000.0

    #11-Габариты
    orow['volpr'] = str(floatx(irow['wight']) * 100) + 'x' + str(floatx(irow['lenght'] )* 100.0) + 'x' + str(floatx(irow['height']) * 100.0)

    #9 - Описание
    orow['descrpr'] = strx(irow['description']).replace('\*/', '<br>')

    return orow


# Copy
i = 1
with open('output.csv', 'wb') as writecsvfile:
    out_fieldnames = ['index', 'iddm', 'art', 'catid', 'namepr', 'brandpr', 'model', 'colorpr', 'descrpr', 'sexpr', 'volpr', 'wightpr', 'lenghtpr', 'heightpr', 'weightpr', 'structpr', 'countrypr', 'agefrompr', 'agetopr', 'kgt', 'recomm', 'tags', 'rating', 'seasonpr', 'price', 'pricevol']
    writer = csv.DictWriter(writecsvfile, fieldnames=out_fieldnames, delimiter=',', quotechar='"')
    hrow = {'index': 'н/п', 'iddm': 'Уникальный ID ДМ', 'art': 'Артикул', 'catid': 'Категория', 'namepr': 'Наименование', 'brandpr': 'Бренд', 'colorpr': 'Цвет', 'descrpr': 'Описание', 'sexpr': 'Пол', 'volpr': 'Габариты', 'wightpr': 'Ширина см', 'lenghtpr':'Длина см', 'heightpr':'Высота см', 'weightpr': 'Вес', 'structpr':'Материал', 'countrypr':'Страна производитель', 'agefrompr':'Возрастная группа от', 'agetopr':'Возрастная группа до', 'kgt':'КГТ', 'recomm':'Рекомендуем?', 'tags':'теги, через запятую', 'rating':'Рейтинг', 'seasonpr':'Коллекция', 'price':'Цена', 'pricevol':'Ценовой уровень'}
    writer.writerow(hrow)

    with open('input.csv', 'rb') as readcsvfile:
        in_fieldnames = ['iddm', 'art', 'name', 'brandid','brandname', 'countryid', 'countryname', '1', '2', 'weight', 'weightid','lenght', 'wight', 'height', 'volume', 'createdate', 'sex', 'productstr', 'groupproduct', 'namegroupproduct', 'colorid', 'colorname', 'sizeid', 'sizename', 'agefromid', 'agefromname', 'agetoid', 'agetoname', 'season', 'seasonyear', 'description']
        reader = csv.DictReader(readcsvfile, fieldnames=in_fieldnames, delimiter='|')
        for irow in reader:
            if unicode(irow['iddm'])[-3:] == u'001':
                orow = get_output_row(irow, True)
                orow['index'] = str(i)
                writer.writerow(orow)
                i += 1

            orow = get_output_row(irow, False)
            orow['index'] = str(i)
            i += 1
            writer.writerow(orow)

