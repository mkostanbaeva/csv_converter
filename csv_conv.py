#coding=utf-8

#from openpyxl import load_workbook
import csv
import requests
#from Tkinter import Tk
#from tkFileDialog import askopenfilename, asksaveasfilename
import re


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
        vals = irow['name'].value.split(u':')
        if is_sup:
            orow['namepr'].value = vals[0].capitalize()
            orow['model'].value = vals[1]
        elif len(vals) >= 3:
            orow['namepr'].value = vals[0].capitalize()
            orow['model'].value = vals[2]
        else:
            orow['model'].value = irow['name'].value

    # 4-Категория
    if irow['groupproduct'].value:
        id_erp = re.sub(regex, '', irow['groupproduct'].value)
        if id_erp in category_dict:
            orow['catid'].value = category_dict[id_erp]
        else:
            orow['catid'].value = irow['groupproduct'].value

    # 6-Бренд
    if irow['brandname'].value:
        brand = re.sub(regex, '', irow['brandname'].value)
        if brand and brand.replace(' ', '').lower() in brands_dict:
            orow['brandpr'].value = brands_dict[brand.replace(' ', '').lower()]
        else:
            orow['brandpr'].value = irow['brandname'].value

    # 10-пол
    if irow['sex'].value:
        if irow['sex'].value.strip().lower() == u'f':
            orow['sexpr'].value = u'ж'
        elif irow['sex'].value.strip().lower() == u'm':
            orow['sexpr'].value = u'м'
        elif irow['sex'].value.strip().lower() == u'u':
            orow['sexpr'].value = ' '
        else:
            orow['sexpr'].value = irow['sex'].value

    # 16-состав товара
    irow['structpr'].value = irow['productstr'].value

    # 18-Возраст от
    if irow['agefromid'].value:
        if irow['agefromid'].value.strip().lower() == u'0M+':
            orow['agefrompr'].value = u'0'
        elif irow['agefromid'].value.strip().lower() == u'3M+':
            orow['agefrompr'].value = u'3'
        elif irow['agefromid'].value.strip().lower() == u'6M+':
            orow['agefrompr'].value = u'6'
        elif irow['agefronid'].value.strip().lower() == u'9M+':
            orow['agefrompr'].value = u'9'
        elif irow['agefromid'].value.strip().lower() == u'1,5Y+':
            orow['agefrompr'].value = u'18'
        elif irow['agefromid'].value.strip().lower() == u'10Y+':
            orow['agefrompr'].value = u'120'
        elif irow['agefronid'].value.strip().lower() == u'11Y+':
            orow['agefrompr'].value = u'132'
        elif irow['agefromid'].value.strip().lower() == u'12Y+':
            orow['agefrompr'].value = u'144'
        elif irow['agefromid'].value.strip().lower() == u'1Y+':
            orow['agefrompr'].value = u'12'
        elif irow['agefromid'].value.strip().lower() == u'2Y+':
            orow['agefrompr'].value = u'24'
        elif irow['agefromid'].value.strip().lower() == u'3Y+':
            orow['agefrompr'].value = u'36'
        elif irow['agefromid'].value.strip().lower() == u'5Y+':
            orow['agefrompr'].value = u'60'
        elif irow['agefromid'].value.strip().lower() == u'6Y+':
            orow['agefrompr'].value = u'72'
        elif irow['agefromid'].value.strip().lower() == u'7Y+':
            orow['agefrompr'].value = u'84'
        elif irow['agefromid'].value.strip().lower() == u'8Y+':
            orow['agefrompr'].value = u'96'
        elif irow['agefromid'].value.strip().lower() == u'9Y+':
            orow['agefrompr'].value = u'108'
        elif irow['agefromid'].value.strip().lower() == u'ADULT':
            orow['agefrompr'].value = u' '
        else:
            orow['agefrompr'].value = irow['agefromid'].value

    # 19-Возраст до
    if irow['agetoid'].value:
        if irow['agetoid'].value.strip().lower() == u'3M-':
            orow['agetopr'].value = u'3'
        elif irow['agetoid'].value.strip().lower() == u'6M-':
            orow['agetopr'].value = u'6'
        elif irow['agetoid'].value.strip().lower() == u'9M-':
            orow['agetopr'].value = u'9'
        elif irow['agetoid'].value.strip().lower() == u'1,5Y-':
            orow['agetopr'].value = u'18'
        elif irow['agetoid'].value.strip().lower() == u'10Y-':
            orow['agetopr'].value = u'120'
        elif irow['agetoid'].value.strip().lower() == u'11Y-':
            orow['agetopr'].value = u'132'
        elif irow['agetoid'].value.strip().lower() == u'12Y-':
            orow['agetopr'].value = u'144'
        elif irow['agetoid'].value.strip().lower() == u'1Y-':
            orow['agetopr'].value = u'12'
        elif irow['agetoid'].value.strip().lower() == u'2Y-':
            orow['agetopr'].value = u'24'
        elif irow['agetoid'].value.strip().lower() == u'3Y-':
            orow['agetopr'].value = u'36'
        elif irow['agetoid'].value.strip().lower() == u'4Y-':
            orow['agetopr'].value = u'48'
        elif irow['agetoid'].value.strip().lower() == u'5Y-':
            orow['agetopr'].value = u'60'
        elif irow['agetoid'].value.strip().lower() == u'6Y-':
            orow['agetopr'].value = u'72'
        elif irow['agetoid'].value.strip().lower() == u'7Y-':
            orow['agetopr'].value = u'84'
        elif irow['agetoid'].value.strip().lower() == u'8Y-':
            orow['agetopr'].value = u'96'
        elif irow['agetoid'].value.strip().lower() == u'9Y-':
            orow['agetopr'].value = u'108'
        else:
            orow['agetopr'].value = irow['agetoid'].value

    #24-Коллекция
    orow['seasonpr'].value = irow['season'].value + irow['seasonyear'].value

    #17 - Страна происхождения
    orow['countrypr'].value = irow['countryname'].value

    #12 - Ширина
    orow['wightpr'].value = irow['wight'].value * 100

    #14 - Высота
    orow['heightpr'].value = irow['height'].value * 100

    #13 - Длина
    orow['lenghtpr'] = irow['lenght'].value * 100

    #15 - Вес
    orow['weightpr'].value = irow['weight'].value * 1000

    #11-Габариты
    orow['volpr'].value = unicode('{0:.0f}'.format(float(irow['wight'].value) * 100)) + u'x' + unicode('{0:.0f}'.format(float(irow['lenght'].value )* 100.0)) + u'x' + unicode('{0:.0f}'.format(float(irow['height'].value) * 100.0))

    #9 - Описание
    orow['descrpr'].value = irow['description'].value

    return orow


# Copy
i = 1
with open('output.csv', 'wb') as writecsvfile:
    out_fieldnames = ['index', 'iddm', 'art', 'catid', 'namepr', 'brandpr', 'model', 'colorpr', 'descrpr', 'sexpr', 'volpr', 'wightpr', 'lenghtpr', 'heightpr', 'wightpr', 'structpr', 'countrypr', 'agefronpr', 'agetopr', 'kgt', 'recomm', 'tags', 'rating', 'seasonpr', 'price', 'pricevol']
    writer = csv.DictWriter(writecsvfile, fieldnames=out_fieldnames, delimiter=',')
    writer.writeheader()

    with open('input.csv', 'rb') as readcsvfile:
        in_fieldnames = ['iddm', 'art', 'name', 'brandid','brandname', 'countryid', 'countryname', '1', '2', 'weight', 'weightid','lenght', 'wight', 'height', 'volume', 'createdate', 'sex', 'productstr', 'groupproduct', 'namegroupproduct', 'colorid', 'colorname', 'sizeid', 'sizename', 'agefromid', 'agefromname', 'agetoid', 'agetoname', 'season', 'seasonyear', 'description']
        reader = csv.DictReader(readcsvfile, fieldnames=in_fieldnames, delimiter=',')
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


orow['index'].value = u'н/п'
orow['iddm'].value = u'уникальный ID ДМ'
orow['art'].value = u'Артикул'
orow['catid'].value = u'Категория'
orow['namepr'].value = u'Наименование'
orow['brandpr'].value = u'Бренд'
orow['model'].value = u'Модель'
orow['colorpr'].value = u'Цвет'
orow['descrpr'].value = u'Описание'
orow['sexpr'].value = u'Пол'
orow['volpr'].value = u'Габариты'
orow['wightpr'].value = u'Ширина см'
orow['lenghtpr'].value = u'Длина см'
orow['heightpr'].value = u'Высота см'
orow['weightpr'].value = u'Вес'
orow['structpr'].value = u'Материал'
orow['countrypr'].value = u'Страна производитель'
orow['agefrompr'].value = u'Возрастная группа от'
orow['agetopr'].value = u'Возрастная группа до'
orow['kgt'].value = u'КГТ'
orow['recomm'].value = u'Рекомендуем?'
orow['tags'].value = u'теги, через запятую'
orow['rating'].value = u'Рейтинг'
orow['seasonpr'].value = u'Коллекция'
orow['price'].value = u'Цена'
orow['pricevol'].value = u'Ценовой уровень'

# Save the file
#wb.save("result.xlsx")
