import pandas as pd
import numpy as np
import \
    xlrd  # чтение эксель файлов - для чтения эксель файлов из 1с нужна обязательно версия 1.2, версия 2 не читает
# эксель из 1С, openpyxl так же не читает файлы из 1С
from pathlib import Path
import os

# Шаблон
template_paid_free = pd.read_excel(r'шаблон_заявки_по_жк.xlsx', sheet_name='платные_то')
v = list(template_paid_free['#'][:])
template_paid_free.index = v
template_paid_free = template_paid_free.drop('#', axis=1)

template_cleaning = pd.read_excel(r'шаблон_заявки_по_жк.xlsx', sheet_name='клининг')
v = list(template_cleaning['#'][:])
template_cleaning.index = v
template_cleaning = template_cleaning.drop('#', axis=1)

root_dir = input('Укажите полный путь к папке, в которой выгрузены файлы из 1С, затем нажмите Enter: ')
root_dir = root_dir.strip()
root_dir = root_dir.rstrip()

root_dir2 = input('Укажите полный путь к папке, где будет сохранена итоговвая таблица, затем нажмите Enter: ')
root_dir2 = root_dir2.strip()
root_dir2 = root_dir2.rstrip()

# Переходим в директорию, где лежат выгрузки из 1С
os.chdir(root_dir)

# Подгружаем и обрабатываем данные из 1С
vc = xlrd.open_workbook(r'вк.xlsx')
vc = pd.read_excel(vc)
ritc = xlrd.open_workbook(r'ритц.xlsx')
ritc = pd.read_excel(ritc)
gs = xlrd.open_workbook(r'гс.xlsx')
gs = pd.read_excel(gs)

complete_table = pd.concat([vc, ritc, gs], ignore_index=False)

'''Платные заявки'''
# Сортировка
mask_paid_applications = complete_table['Вид заявки'].isin(['Платная'])
# сводная таблица
pivot_paid_applications = pd.pivot_table(complete_table[mask_paid_applications], index = 'Жилой комплекс', columns = 'Статус выполнения', values = 'Дата',
                                    aggfunc = np.count_nonzero, fill_value = 0, margins = True)
# ищем кол-во поступивших заявок
pivot_paid_applications['Выполнено заявок'] = pivot_paid_applications[['Выполнено', 'Закрыта', 'Контроль']].sum(axis=1)
# впр на три необходимых столбца
paid_applications = pd.merge(template_paid_free, pivot_paid_applications, how = 'left', left_on = 'ЖК', right_on = 'Жилой комплекс')[['ЖК', 'Управление', 'All', 'Выполнено заявок', 'Новая заявка']].rename(columns = {'All' : 'Поступило заявок'})
# заполняем пропуски 0
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
def summ_region (excel, region_name):
    # excel - название DataFrame (DF)
    # region_name - название управление по которому формируем итог по статусам
    # Создаем пустой список для индексов строк относящихся к заданному управлению
    index = []
    # Перебираем значения от 0 до размера DF по строкам
    for reg in range (len(excel)):
        # Если значение в столбце Управление равно заданному
        if excel.iat[reg,1] == region_name:
            # то записываем индекс строки в список
            index.append(reg)
    # Суммиуем значения одного управления по столбцам и создаем из этих значений список
    region_sum = list(excel.iloc[index[1]:index[-1]:].sum(axis=0))
    # Так как у нас идеи суммирование по всем столбцам, то и столбец с перечнем ЖК тоже суммируется
    # и мы плучаем длинное-длинное название содиненное из всех ЖК
    # Здесь мы переименовываем первые два значения в корректные
    region_sum[0] = 'Итого:'
    region_sum[1] = ''
    # Записываем суммы в соответсвующую строку
    excel.loc[index[-1]] = region_sum
    # Результат ф-и таблица с суммой по указанному управлениую
    return excel
paid_applications = summ_region(paid_applications, 'Бизнес-класс')
paid_applications = summ_region(paid_applications, 'Восток')
paid_applications = summ_region(paid_applications, 'Запад')
paid_applications = summ_region(paid_applications, 'Прочее')
# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
paid_applications = paid_applications.drop('Управление', axis=1)

'''Бесплатные заявки'''
# Сортировка
mask_free_applications = complete_table['Вид заявки'].isin(['Бесплатная'])
# сводная таблица
pivot_free_applications = pd.pivot_table(complete_table[mask_free_applications], index = 'Жилой комплекс', columns = 'Статус выполнения', values = 'Дата',
                                    aggfunc = np.count_nonzero, fill_value = 0, margins = True)
# впр на два необходимых столбца
free_applications = pd.merge(template_paid_free, pivot_free_applications, how = 'left', left_on = 'ЖК', right_on = 'Жилой комплекс')[['ЖК', 'Управление', 'All', 'Новая заявка']].rename(columns = {'All' : 'Количество'})
# заполняем пропуски 0
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
free_applications = summ_region(free_applications, 'Прочее')
# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
free_applications = free_applications.drop('Управление', axis=1)

'''Клининг'''
# Сортировка
mask_cleaning = complete_table['Вид работ'].isin(['Клининг'])
# сводная таблица
pivot_cleaning = pd.pivot_table(complete_table[mask_cleaning], index = 'Жилой комплекс', columns = 'Статус выполнения', values = 'Дата',
                                    aggfunc = np.count_nonzero, fill_value = 0, margins = True)
# впр на два необходимых столбца
cleaning = pd.merge(template_cleaning, pivot_cleaning, how = 'left', left_on = 'ЖК', right_on = 'Жилой комплекс')[['ЖК', 'Управление', 'All', 'Новая заявка']].rename(columns = {'All' : 'Количество'})
# заменяем пропуски на 0
cleaning = cleaning.fillna(0)
# Проверяем дублирование объектов (убираем тестовые заявки)
cleaning = dubl(cleaning)
# Добавляем итоговую сумму по всем объектам - последняя строка
cleaning = cleaning.append(cleaning.sum(axis=0), ignore_index=True)
cleaning.at[len(cleaning) - 1, 'ЖК'] = 'Итого'
cleaning.at[len(cleaning) - 1, 'Управление'] = ''
# Считаем суммы по управлениям
cleaning = summ_region(cleaning, 'Бизнес-класс')
cleaning = summ_region(cleaning, 'Восток')
cleaning = summ_region(cleaning, 'Запад')
cleaning = summ_region(cleaning, 'Прочее')
# Удаляем признак "Управление", т.к. в итоговой таблице он нам не нужен
cleaning = cleaning.drop('Управление', axis=1)

# Переходим в директорию, куда необходимо сохранить новый файл
os.chdir(root_dir2)

# Создание excel-файла и запись в него итоговой таблицы
writer = pd.ExcelWriter('заявки_жк_месяц.xlsx')
paid_applications.to_excel(writer, 'платные')
free_applications.to_excel(writer, 'бесплатные')
cleaning.to_excel(writer, 'клининг')
writer.save()