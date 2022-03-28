import pandas as pd
import numpy as np
import \
    xlrd  # чтение эксель файлов - для чтения эксель файлов из 1с нужна обязательно версия 1.2, версия 2 не читает
# эксель из 1С, openpyxl так же не читает файлы из 1С
from xlrd import Book
from pathlib import Path
import os

''' Подгружаем формы отчетов '''
# Шаблон для платных и то заявок
template_paid_free = xlrd.open_workbook(r"шаблон_платные_то.xlsx")
template_paid_free = pd.read_excel(template_paid_free)
v = list(template_paid_free['#'][:])
template_paid_free.index = v
template_paid_free = template_paid_free.drop('#', axis=1)
# Шаблон для заявок по клинингу
template_clean = xlrd.open_workbook(r"шаблон_клининг.xlsx")
template_clean = pd.read_excel(template_clean)
v = list(template_clean['#'][:])
template_clean.index = v
template_clean = template_clean.drop('#', axis=1)

''' Работа с директориями '''
root_dir = input('Укажите полный путь к папке, в которой выгрузены файлы из 1С, затем нажмите Enter: ')
root_dir = root_dir.strip()
root_dir = root_dir.rstrip()
root_dir2 = input('Укажите полный путь к папке, где будет сохранена итоговвая таблица, затем нажмите Enter: ')
root_dir2 = root_dir2.strip()
root_dir2 = root_dir2.rstrip()
# Переходим в директорию, где лежат выгрузки из 1С
os.chdir(root_dir)

'''Подгружаем данные выгруженные из 1С '''
# Платные заявки
paid_vc_1c = xlrd.open_workbook(r'платные_вк.xlsx')
paid_vc_1c = pd.read_excel(paid_vc_1c)
paid_ritc_1c = xlrd.open_workbook(r'платные_ритц.xlsx')
paid_ritc_1c = pd.read_excel(paid_ritc_1c)
paid_gs_1c = xlrd.open_workbook(r'платные_гс.xlsx')
paid_gs_1c = pd.read_excel(paid_gs_1c)
# Заявки ТО
free_vc_1c = xlrd.open_workbook(r'то_вк.xlsx')
free_vc_1c = pd.read_excel(free_vc_1c)
free_ritc_1c = xlrd.open_workbook(r'то_ритц.xlsx')
free_ritc_1c = pd.read_excel(free_ritc_1c)
free_gs_1c = xlrd.open_workbook(r'то_гс.xlsx')
free_gs_1c = pd.read_excel(free_gs_1c)
# Заявки по клинингу
clean_vc_1c = xlrd.open_workbook(r'клининг_вк.xlsx')
clean_vc_1c = pd.read_excel(clean_vc_1c)
clean_ritc_1c = xlrd.open_workbook(r'клининг_гс.xlsx')
clean_ritc_1c = pd.read_excel(clean_ritc_1c)
clean_gs_1c = xlrd.open_workbook(r'клининг_ритц.xlsx')
clean_gs_1c = pd.read_excel(clean_gs_1c)

''' Платные заявки '''

paid_vc_1c = paid_vc_1c[:-1]
paid_ritc_1c = paid_ritc_1c[:-1]
paid_gs_1c = paid_gs_1c[:-1]

# ф-я убирает все строки выше основной таблицы - строки расположеные выше ячейки с текстом "В работе"
def v_rabote(excel):
    for row in range(excel.shape[0]):
        for col in range(excel.shape[1]):
            if excel.iat[row, col] == 'В работе':  # row - строка
                return excel.iloc[row:].reset_index(drop=True)
                break
paid_vc_1c = v_rabote(paid_vc_1c)
paid_ritc_1c = v_rabote(paid_ritc_1c)
paid_gs_1c = v_rabote(paid_gs_1c)

# ф-я убирает строку между данными и шапкой таблицы - строка с ячейкой "_Все ЖК"
def _vse_gk(excel):
    for row in range(excel.shape[0]):
        for col in range(excel.shape[1]):
            if excel.iat[row, col] == '_Все ЖК':
                return excel.drop([row], axis=0).reset_index(drop=True)
                break
paid_vc_1c = _vse_gk(paid_vc_1c)
paid_ritc_1c = _vse_gk(paid_ritc_1c)
paid_gs_1c = _vse_gk(paid_gs_1c)

# Делаем первый столбец с ЖК индексом и убираем его
v = list(paid_vc_1c['Unnamed: 0'][:])
paid_vc_1c.index = v
paid_vc_1c = paid_vc_1c.drop('Unnamed: 0', axis=1)
v = list(paid_ritc_1c['Unnamed: 0'][:])
paid_ritc_1c.index = v
paid_ritc_1c = paid_ritc_1c.drop('Unnamed: 0', axis=1)
v = list(paid_gs_1c['Unnamed: 0'][:])
paid_gs_1c.index = v
paid_gs_1c = paid_gs_1c.drop('Unnamed: 0', axis=1)

# Переносим первую строку в названия признаков и убираем нулевую строку
paid_vc_1c.columns = paid_vc_1c.iloc[0]
paid_vc_1c = paid_vc_1c[1:]
paid_ritc_1c.columns = paid_ritc_1c.iloc[0]
paid_ritc_1c = paid_ritc_1c[1:]
paid_gs_1c.columns = paid_gs_1c.iloc[0]
paid_gs_1c = paid_gs_1c[1:]

# Соединаем три таблицы друг под другом
paid_applications_0 = pd.concat([paid_vc_1c, paid_ritc_1c, paid_gs_1c], ignore_index=False)

# ВПР данных в шаблон и пересчитываем индексы, т.к. могут попасться тестовые заявки
# в этом случае идексы задвоятся, будут неправильно чиститься
paid_applications = pd.merge(template_paid_free, paid_applications_0, how='left', left_on='ЖК', right_index=True).reset_index(drop=True)

# Удалили столбцы где нет ни одного значения
paid_applications.dropna(axis='columns', how='all', inplace=True)

# заменяем пропущенные значения на 0
paid_applications = paid_applications.fillna(0)

# Проверяем дублирование объектов (убираем тестовые заявки)
def dubl(excel):
    nan_ = []
    nan_z = '0'
    for n in range(len(excel.columns)):
        nan_.append(nan_z)
    for row1 in range(len(excel) - 1):
        for row2 in range(len(excel) - 1):
            if excel.iat[row1, 0] == excel.iat[row2, 0]:
                if int(excel.iat[row1, len(excel.columns) - 1]) > int(excel.iat[row2, len(excel.columns) - 1]):
                    excel.loc[row2] = nan_
                elif int(excel.iat[row1, len(excel.columns) - 1]) < int(excel.iat[row2, len(excel.columns) - 1]):
                    excel.loc[row1] = nan_
    excel = excel.loc[excel['ЖК'] != '0']
    return excel.reset_index(drop=True)

paid_applications = dubl(paid_applications)

# Добавляем итоговую сумму по всем объектам - последняя строка
paid_applications = paid_applications.append(paid_applications.sum(axis=0), ignore_index=True)
paid_applications.at[len(paid_applications) - 1, 'ЖК'] = 'Итого'
paid_applications.at[len(paid_applications) - 1, 'Управление'] = ''

# Считаем суммы по управлениям
def summ_region(excel, region_name):
    index = []
    for reg in range(len(excel)):
        if excel.iat[reg, 1] == region_name:
            index.append(reg)
    region_sum = list(excel.iloc[index[0] + 1 : index[-1] + 1, :].sum(axis=0))
    region_sum[0] = region_name
    region_sum[1] = region_name
    excel.loc[index[0]] = region_sum
    return excel

paid_applications = summ_region(paid_applications, 'Бизнес-класс')
paid_applications = summ_region(paid_applications, 'Восток')
paid_applications = summ_region(paid_applications, 'Запад')
paid_applications = summ_region(paid_applications, 'Обнинск')
paid_applications = summ_region(paid_applications, 'Прочее')

# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
paid_applications = paid_applications.drop('Управление', axis=1)

# В итоговой таблице последний признак имеет другое название - переименовываем
paid_applications = paid_applications.rename(columns={'Невыполненных в прошлом периоде': 'Всего не выполнено на дату отчета'})



'''Заявки ТО '''
# удалим последние строки
free_vc_1c = free_vc_1c[:-1]
free_ritc_1c = free_ritc_1c[:-1]
free_gs_1c = free_gs_1c[:-1]

# убираем все строки выше основной таблицы - строки расположеные выше ячейки с текстом "В работе"
free_vc_1c = v_rabote(free_vc_1c)
free_ritc_1c = v_rabote(free_ritc_1c)
free_gs_1c = v_rabote(free_gs_1c)

# убираем строку между данными и шапкой таблицы - строка с ячейкой "_Все ЖК"
free_vc_1c = _vse_gk(free_vc_1c)
free_ritc_1c = _vse_gk(free_ritc_1c)
free_gs_1c = _vse_gk(free_gs_1c)

# Делаем первый столбец с ЖК индексом и убираем его
v = list(free_vc_1c['Unnamed: 0'][:])
free_vc_1c.index = v
free_vc_1c = free_vc_1c.drop('Unnamed: 0', axis=1)
v = list(free_ritc_1c['Unnamed: 0'][:])
free_ritc_1c.index = v
free_ritc_1c = free_ritc_1c.drop('Unnamed: 0', axis=1)
v = list(free_gs_1c['Unnamed: 0'][:])
free_gs_1c.index = v
free_gs_1c = free_gs_1c.drop('Unnamed: 0', axis=1)

# Переносим первую строку в названия признаков и убираем нулевую строку
free_vc_1c.columns = free_vc_1c.iloc[0]
free_vc_1c = free_vc_1c[1:]
free_ritc_1c.columns = free_ritc_1c.iloc[0]
free_ritc_1c = free_ritc_1c[1:]
free_gs_1c.columns = free_gs_1c.iloc[0]
free_gs_1c = free_gs_1c[1:]

# Соединаем тра таблицы друг под другом
free_applications_0 = pd.concat([free_vc_1c, free_ritc_1c, free_gs_1c], ignore_index=False)

# ВПР данных в шаблон и пересчитываем индексы, т.к. могут попасться тестовые заявки
# в этом случае идексы задвоятся, будут неправильно чиститься
free_applications = pd.merge(template_paid_free, free_applications_0, how='left', left_on='ЖК', right_index=True).reset_index(drop=True)

# Удалили столбцы где нет ниодного значения
free_applications.dropna(axis='columns', how='all', inplace=True)

# заменяем пропущенные значения на 0
free_applications = free_applications.fillna(0)

# Проверяем дублирование объектов (убираем тестовые заявки)
free_applications = dubl(free_applications)

# Добавляем итоговую сумму по всем объектам - последняя строка
free_applications = free_applications.append(free_applications.sum(axis=0), ignore_index=True)
free_applications.at[len(free_applications) - 1, 'ЖК'] = 'Итого'
free_applications.at[len(free_applications) - 1, 'Управление'] = ''

# Считаем суммы по управлениям
free_applications = summ_region(free_applications, 'Бизнес-класс')
free_applications = summ_region(free_applications, 'Восток')
free_applications = summ_region(free_applications, 'Запад')
free_applications = summ_region(free_applications, 'Обнинск')
free_applications = summ_region(free_applications, 'Прочее')

# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
free_applications = free_applications.drop('Управление', axis=1)

# В итоговой таблице последний признак имеет другое название - переименовываем
free_applications = free_applications.rename(columns={'Невыполненных в прошлом периоде': 'Всего не выполнено на дату отчета'})


''' Заявки по клинингу '''

# Соединяем выгрузки друг под другом
cleaning_0 = pd.concat([clean_vc_1c, clean_ritc_1c, clean_gs_1c], ignore_index=False)

# Делаем сводную таблицу как делаем это в excel
cleaning_0 = pd.pivot_table(cleaning_0, index = 'Жилой комплекс', columns = 'Статус выполнения', values = 'Дата', aggfunc = np.count_nonzero, fill_value = 0, margins = True)

# Делаем ВПР данных из своднй таблицы в шаблон
cleaning = pd.merge(template_clean, cleaning_0, how = 'left', left_on = 'ЖК', right_index = True)

# Добавляем строку внизу с подсчетом значений по столбцам, так создастся итоговый столбец с суммами по строкам
cleaning.dropna(axis=1,how='all', inplace=True)
cleaning = cleaning.append(cleaning.sum(axis=0), ignore_index=True)
cleaning.at[len(cleaning)-1, 'ЖК'] = 'Итого'
cleaning.at[len(cleaning)-1, 'Управление'] = ''

# Применяем ф-ю ко всем управлениям
def summ_region_clean(excel, region_name):
    index = []
    for reg in range(len(excel)):
        if excel.iat[reg, 1] == region_name:
            index.append(reg)
    region_sum = list(excel.iloc[index[0] + 1 : index[-1] + 1, :].sum(axis=0))
    region_sum[0] = 'Итого:'
    region_sum[1] = ''
    excel.loc[index[-1]] = region_sum
    return excel
cleaning = summ_region_clean(cleaning, 'Бизнес-класс')
cleaning = summ_region_clean(cleaning, 'Восток')
cleaning = summ_region_clean(cleaning, 'Запад')
cleaning = summ_region_clean(cleaning, 'Прочее')

# Переименовываем последний столбец
cleaning = cleaning.rename(columns = {'All' : 'Итог'})

# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
cleaning = cleaning.drop('Управление', axis=1)

# заменяем пропущенные значения на 0
cleaning = cleaning.fillna(0)



''' Записываем таблицы '''

# Переходим в директорию куда необходимо сохранить готовую таблицу
os.chdir(root_dir2)

# Создание excel-файла и запись в него итоговой таблицы
writer = pd.ExcelWriter('заявки_жк_неделя.xlsx')
paid_applications.to_excel(writer, 'платные')
free_applications.to_excel(writer, 'бесплатные')
cleaning.to_excel(writer, 'клининг')
writer.save()