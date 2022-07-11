from selenium import webdriver
import time
import os
import pandas as pd
import gspread
import tempfile
import shutil


def browser(base_path, vhod):
    login = input('Введите логин цифровой платформы:\n')
    password = input('Введите пароль цифровой платформы:\n')
    platform_url = input('Введите url админской части:\n')
    student_url = input('Введите url студенческой части:\n')
    
    
    
    tempdir = tempfile.mkdtemp(prefix="AutoUpdateSZ-")
    
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument('--disable-gpu')
    options.add_experimental_option("prefs",
                                    {"download.default_directory": tempdir, "download.prompt_for_download": False,
                                     "download.directory_upgrade": True, "safebrowsing.enabled": True})
    web = webdriver.Chrome('./driver/chromedriver.exe', options=options)
    web.get(f'{platform_url}')
    web.find_element_by_xpath('//*[@id="id_username"]').send_keys(f'{login}')
    web.find_element_by_xpath('//*[@id="id_password"]').send_keys(f'{password}')
    web.find_element_by_xpath('//*[@id="login-form"]/div[4]/input').click()
    web.get(f'{student_url}')
    web.find_element_by_xpath('//*[@id ="action-toggle"]').click()
    web.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/span[3]/a').click()
    web.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/label/select/option[2]').click()
    web.find_element_by_xpath('//*[@id="changelist-form"]/div[1]/button').click()
    time.sleep(3)
    web.quit()
    
    os.chdir(tempdir)
    check_name = os.listdir()[0]
    os.rename(check_name, "site.xlsx")
    check_name = os.listdir()[0]
    
    if vhod == 4:
        df1 = pd.read_excel(check_name, sheet_name='Sheet1',
                            usecols=['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'СНИЛС', 'Категория слушателя',
                                     'Компетенция', 'Выбранное место обучения', 'Статус заявки на обучение', 'Email'],
                            dtype=object)
    
    if vhod == 5:
        df1 = pd.read_excel(check_name, sheet_name='Sheet1', usecols=['СНИЛС', 'Email'], dtype=object)
    
    else:
        df1 = pd.read_excel(check_name, sheet_name='Sheet1',
                            usecols=['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'СНИЛС', 'Email', 'Телефон',
                                     'Город проживания', 'Регион проживания', 'Категория слушателя', 'Компетенция',
                                     'Выбранное место обучения', 'Адрес выбранного место обучения',
                                     'Статус заявки на обучение', 'Группа', 'Тип договора', 'Дата начала обучения',
                                     'Дата окончания обучения', 'Занятость по итогам обучения'], dtype=object)
    
    os.remove(check_name)
    os.chdir(base_path)
    shutil.rmtree(tempdir)
    return df1


def main():
    while True:
        
        vhod = input('1-Добавить новые в Учет, 2-Сделать сводник, 3-Статусы в WSR, 4-Список для ЦЗН, 5-Обновить ЦЗН')
        
        if not vhod.isnumeric():
            print('Введено некорректное значение! Повторите \n')
            main()
        else:
            break
    
    base_path = os.getcwd()
    
    # Добавление новых заявок в google sheets
    if vhod == 1:
        df1 = browser(base_path, vhod)
        df3 = df1.drop(
            ['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'СНИЛС', 'Телефон', 'Город проживания', 'Регион проживания',
             'Компетенция', 'Выбранное место обучения', 'Адрес выбранного место обучения', 'Статус заявки на обучение',
             'Группа', 'Тип договора', 'Дата начала обучения', 'Дата окончания обучения',
             'Занятость по итогам обучения'], axis=1)
        
        gc = gspread.service_account(filename='api_key.json')
        sh = gc.open("Учёт заявок СЗ")
        
        actual_total = len(df1.index)
        unemp = df1['Занятость по итогам обучения'].isnull().sum().tolist()
        actual_zanyat = actual_total - unemp
        wsheet2 = sh.worksheet('Проверка на новьё')
        wsheet3 = sh.worksheet('Показатели')
        
        wsheet2.update('A2:B', df3.values.tolist())
        wsheet3.update('F3', actual_total)
        wsheet3.update('F20', actual_zanyat)
    
    # Составление сводной информации с сайта и google sheets
    elif vhod == 2:
        df1 = browser(base_path, vhod)
        
        gc = gspread.service_account(filename='api_key.json')
        sh = gc.open("Учёт заявок СЗ")
        
        # Получение всей информации с google sheets
        wsheet = sh.worksheet('Учет')
        data = wsheet.get_all_values()
        headers = data.pop(0)
        
        df2 = pd.DataFrame(data, columns=headers)
        df2 = df2.drop(['Категория слушателя', 'СОПД', 'ПАСПОРТ с пропиской!', 'Где прописка?', 'СНИЛС(от 02.07.21)',
                        'Если меняла фамилию, подтверждающий документ', 'ИЩУЩИЙ', 'БЕЗРАБ (справка/выписка)',
                        'копия трудовой', 'Справка ПРЕДПЕНС', 'ПОДТВЕРЖДЕНИЕ ДЕКРЕТА/справка не ИП', 'Извещение ПФР',
                        'Св-во о рождении ребенка', 'Комментарий', 'Статус последнего прозвона', 'Кто звонил?',
                        'Статусы ВСР'], axis=1)
        df2['ID'] = pd.to_numeric(df2['ID'])
        
        data = pd.merge(df1, df2, left_on='Email', right_on='Email', how='left')
        cols = data.columns.tolist()
        cols = cols[-4:] + cols[:-4]
        data = data[cols]
        data = data.sort_values(by=['ID'])
        data.to_excel('C:/Users/COPP/Desktop/work/СЗ/Свод/Новый_свод.xlsx', index=None)

    # Перенос статусов с цифровой платформы в google sheets
    elif vhod == 3:
        df1 = browser(base_path, vhod)
        gc = gspread.service_account(filename='api_key.json')
        
        # Получение всей информации с google sheets
        sh = gc.open("Учёт заявок СЗ")
        wsheet = sh.worksheet('Учет')
        data = wsheet.get_all_values()
        headers = data.pop(0)
        
        df2 = pd.DataFrame(data, columns=headers)
        df2 = df2.drop(['Категория слушателя', 'СОПД', 'ПАСПОРТ с пропиской!', 'Где прописка?', 'СНИЛС(от 02.07.21)',
                        'Если меняла фамилию, подтверждающий документ', 'ИЩУЩИЙ', 'БЕЗРАБ (справка/выписка)',
                        'копия трудовой', 'Справка ПРЕДПЕНС', 'ПОДТВЕРЖДЕНИЕ ДЕКРЕТА/справка не ИП', 'Извещение ПФР',
                        'Св-во о рождении ребенка', 'Комментарий', 'Статус последнего прозвона', 'Кто звонил?', 'ID',
                        'Статус', 'ДИПЛОМ', 'Свежий ЦЗН', 'Статусы ВСР'], axis=1)
        
        df4 = df1.drop(
            ['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'СНИЛС', 'Телефон', 'Город проживания', 'Регион проживания',
             'Компетенция', 'Категория слушателя', 'Выбранное место обучения', 'Адрес выбранного место обучения',
             'Группа', 'Тип договора', 'Дата начала обучения', 'Дата окончания обучения',
             'Занятость по итогам обучения'], axis=1)
        
        df4 = pd.merge(df2, df4, left_on='Email', right_on='Email', how='left')
        df4 = df4.drop(['Email'], axis=1)
        df4.fillna('Отменен/переведен', inplace=True)
        
        wsheet.update('V2:V', df4.values.tolist())
    
    
    # Составление списка для Центра занятости населения
    elif vhod == 4:
        df1 = browser(base_path, vhod)
        
        # Получение всей информации с google sheets
        gc = gspread.service_account(filename='api_key.json')
        sh = gc.open("Учёт заявок СЗ")
        wsheet = sh.worksheet('Учет')
        data = wsheet.get_all_values()
        headers = data.pop(0)
        df2 = pd.DataFrame(data, columns=headers)
        df2 = df2.drop(['Категория слушателя', 'СОПД', 'ПАСПОРТ с пропиской!', 'СНИЛС(от 02.07.21)',
                        'Если меняла фамилию, подтверждающий документ', 'ИЩУЩИЙ', 'БЕЗРАБ (справка/выписка)',
                        'копия трудовой', 'Справка ПРЕДПЕНС', 'ПОДТВЕРЖДЕНИЕ ДЕКРЕТА/справка не ИП', 'Извещение ПФР',
                        'Св-во о рождении ребенка', 'Комментарий', 'Статус последнего прозвона', 'Кто звонил?', 'ID',
                        'Статус', 'ДИПЛОМ', 'Статусы ВСР'], axis=1)
        df1 = df1.drop(['Телефон', 'Город проживания', 'Регион проживания', 'Адрес выбранного место обучения', 'Группа',
                        'Тип договора', 'Дата начала обучения', 'Дата окончания обучения',
                        'Занятость по итогам обучения'], axis=1)
        
        df1['Фамилия'] = df1['Фамилия'].str.upper()
        df1['Имя'] = df1['Имя'].str.upper()
        df1['Отчество'] = df1['Отчество'].str.upper()
        
        df3 = pd.read_excel('C:/Users/COPP/Desktop/work/СЗ/Зачисленные в ЦЗН/Проверка.xlsx', sheet_name='Sheet1',
                            dtype=object)
        df3['Фамилия'] = df3['Фамилия'].str.upper()
        df3['Имя'] = df3['Имя'].str.upper()
        df3['Отчество'] = df3['Отчество'].str.upper()
        
        data1 = pd.merge(df1, df2, left_on=['Email'], right_on=['Email'], how='left')
        data2 = pd.merge(df3, data1, left_on=['Фамилия', 'Имя', 'Отчество'], right_on=['Фамилия', 'Имя', 'Отчество'],
                         how='left')
        data2.to_excel('C:/Users/COPP/Desktop/work/СЗ/Зачисленные в ЦЗН/Результат.xlsx', index=None)
    
    elif vhod == 5:
        x = input('Введи дату выгрузки! В формате ДД.ММ.ГГ\n')
        
        # Файл с информацией от Центра занятости населения
        df1 = pd.read_excel(f'C:/Users/COPP/Desktop/work/СЗ/Выгрузки ЦЗН/{x}.xls', dtype=object)
        df1.drop(0, inplace=True)
        df1 = df1.drop(df1.columns[[0, 1, 2, 3, 5, 6, 8, 9, 10]], axis=1)
        df1.columns = ['СНИЛС', 'Статус']
        
        # Получение информации с цифровой платформы
        df2 = browser(base_path, vhod)
        
        # Получение информации из google sheets
        gc = gspread.service_account(filename='api_key.json')
        sh = gc.open("Учёт заявок СЗ")
        wsheet = sh.worksheet('Учет')
        data = wsheet.get_all_values()
        headers = data.pop(0)
        
        df3 = pd.DataFrame(data, columns=headers)
        df3 = df3.drop(['Категория слушателя', 'СОПД', 'ПАСПОРТ с пропиской!', 'СНИЛС(от 02.07.21)',
                        'Если меняла фамилию, подтверждающий документ', 'ИЩУЩИЙ', 'БЕЗРАБ (справка/выписка)',
                        'копия трудовой', 'Справка ПРЕДПЕНС', 'ПОДТВЕРЖДЕНИЕ ДЕКРЕТА/справка не ИП', 'Извещение ПФР',
                        'Св-во о рождении ребенка', 'Комментарий', 'Статус последнего прозвона', 'Кто звонил?', 'ID',
                        'Статус', 'ДИПЛОМ', 'Свежий ЦЗН', 'Где прописка?', 'Статусы ВСР'], axis=1)
        
        # Формирование итога
        data1 = pd.merge(df1, df2, left_on=['СНИЛС'], right_on=['СНИЛС'], how='left')
        data2 = pd.merge(data1, df3, left_on=['Email'], right_on=['Email'], how='left')
        data2 = data2.drop_duplicates(subset=['СНИЛС'], keep='first')
        data2 = data2.drop(['СНИЛС'], axis=1)
        data3 = pd.merge(df3, data2, left_on=['Email'], right_on=['Email'], how='left')
        data3.fillna('0', inplace=True)
        data3 = data3.drop(['Email'], axis=1)
        wsheet.update('S2:S', data3.values.tolist())
    
    else:
        print('Промахнулся...')
        main()


if __name__ == "__main__":
    main()
