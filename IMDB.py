import xlrd
import json
import urllib.request
import os
import progressbar
import time
path = os.path.abspath(os.path.dirname(__file__)) + '\kpvotes.xlsx'
FILMS = xlrd.open_workbook(path)
URL = 'http://www.omdbapi.com/?'


def XlsToDict(xls):
    """
    Transfer data from Excel files to list of dictionaries
    """
    lst = list()
    SHEET = xls.sheet_by_index(0)                   # получаем первый лист эксель файла
    for i in range(2, SHEET.nrows):                 # от i = 2 до количества строк
        if SHEET.cell(i, 1).value != '':            # если ячейка с названием пустая, не берем этот фильм
            lst.append({                            # добавляем в конец листа данные о фильме
                'title': str(SHEET.cell(i, 1).value),
                'year': str(SHEET.cell(i, 2).value)[:4],
                'rate': SHEET.cell(i, 7).value,
                'id': '',
            })
    return lst


def GetId(lst):
    bar = progressbar.ProgressBar(maxval=len(lst)).start()  # Создаём новый прогрессбар
    for i in range(len(lst)):
        bar.update(i)
        title = lst[i]['title'].replace(' ', '+')
        year = lst[i]['year']
        req_url = '{}t={}&y={}&r=json'.format(
            URL,
            title,
            year
        )
        data = urllib.request.urlopen(req_url)
        dic = json.loads(data.read().decode(data.info().get_param('charset') or 'utf-8'))
        if dic['Response'] == 'True':
            lst[i]['id'] = dic['imdbID']
        else:
            continue
    bar.finish()  # Заканчиваем обновлять -- далее можно


def printJ(file):
    print(json.dumps(file, indent=4, sort_keys=True))

movielist = XlsToDict(FILMS)
GetId(movielist)
while True:
    print('Type nothing to exit')
    i = int(input(': '))
    if i == -1:
        break
    elif i == '':
        print('type some number')
    else:
        printJ(movielist[i])


# print(movielist)
# input('Press any button...')
