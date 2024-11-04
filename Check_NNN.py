# openpyxl xlsxwriter xlrd pandas
from chromebrowser import ChromeBrowser
from yandexbrowser import YaBrowser
from msedgebrowser import MsEdge
from firefoxbrowser import FireFox
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import os
import pdfplumber
import pandas
import time


def find_pdf(path, fname):
    print(fname)
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.startswith(f'ul-{fname}'):
                fpath = root + '/' + str(file)
                print(f'документ {fpath}')
                return fpath


def find_okved(path):
    pdf = pdfplumber.open(path)
    for table_page in pdf.pages:
        for table in table_page.extract_tables():
            for string in table:
                for cell in string:
                    if cell == 'Код и наименование вида деятельности':
                        print(string[string.index(cell) + 1])
                        okved = (string[string.index(cell) + 1].split()[0])
                        pdf.close()
                        # os.remove(pdf_path)
                        return okved


def find_kpp(path, filial, check_remove):
    check_filial = 1
    pdf = pdfplumber.open(path)
    for table_page in pdf.pages:
        for table in table_page.extract_tables():
            for string in table:
                for cell in string:
                    if cell == 'КПП юридического лица' and filial == 0:
                        print(string[string.index(cell) + 1])
                        kpp = int(string[string.index(cell) + 1].split()[0])
                        pdf.close()
                        if check_remove:
                            os.remove(pdf_path)
                        return kpp
                    if cell == 'Сведения об учете в налоговом органе по\nместу нахождения филиала':
                        print('сведения по филиалу найдены')
                        if check_filial == filial:
                            print(string[string.index(cell) + 1])
                            kpp = int(string[string.index(cell) + 1].split()[2])
                            pdf.close()
                            if check_remove:
                                os.remove(pdf_path)
                            return kpp
                        check_filial += 1


time_start = time.time()
# имя файла откуда берем сведения, файл лежит в папке documents
# документ предварительно подготовлен, первая строка - это наименование столбцов с данными,
# лишних строк сверху быть не должно
# file_name = 'Реестр_Алтайский край.xlsx'
file_name = 'short.xlsx'

# номера столбцов для считывания данных, отсчет начинается с Нуля
cols = [2, 4, 5, 9]

# получаем данные с документа с выбранных колонок
data = pandas.read_excel(f'documents/{file_name}', usecols=cols)
"""
Наименование столбцов (первая строка в документе)
['ИНН\n(10 знаков, без пробелов)']
['КПП\n(9 знаков, без пробелов)']
['ОКВЭД']
['ОКТМО 11']
"""
# формируем списки из исходных данных для проверки
inn_list = []
kpp_isxod = []
okved_isxod = []
oktmo_isxod = []

for row_num in range(len(data)):
    inn_list.append(data['ИНН\n(10 знаков, без пробелов)'][row_num])
    kpp_isxod.append(data['КПП\n(9 знаков, без пробелов)'][row_num])
    okved_isxod.append(data['ОКВЭД'][row_num])
    oktmo_isxod.append(data['ОКТМО 11'][row_num])

# часть кода для получения данных с ФНС и Росстата
kpp_get = []
okved_get = []
oktmo_get = []
ogrn_get = []

ibrowser = YaBrowser()

for count in range(len(inn_list)):
    ibrowser.get_url("https://egrul.nalog.ru/index.html")
    ibrowser.send_keys_by_xpath('//input[@name="query"]', str(inn_list[count]))
    ibrowser.click_by_xpath('//button[@id="btnSearch"]')
    sleep(1.5)
    infolist = ibrowser.driver.find_element(By.XPATH, '//div[@class="res-text"]').text.split()
    ogrn_get.append(infolist[infolist.index('ОГРН:') + 1][:-1])

    if count == 0:
        ibrowser.click_by_xpath('//button[@class="btn-with-icon btn-excerpt op-excerpt"]')
    if count != 0:
        if inn_list[count] != inn_list[count - 1]:
            ibrowser.click_by_xpath('//button[@class="btn-with-icon btn-excerpt op-excerpt"]')
            # если комп не тянет или очень слабый интернет, то увеличивать задержку
    sleep(2)
    print(f'данные получены для {inn_list[count]}')
sleep(2)
"""
filial = 0
for count in range(len(inn_list)):
    print(f'ОКТМО для ИНН - {inn_list[count]}')
    if count != 0:
        if inn_list[count] == inn_list[count - 1]:
            filial += 1
        else:
            filial = 0
    ibrowser.get_url("https://websbor.rosstat.gov.ru/online/info")
    ibrowser.send_keys_by_xpath('//input[@id="inn"]', str(inn_list[count]))
    ibrowser.click_by_xpath('//button[@class="mat-focus-indicator mat-flat-button mat-button-base mat-primary"]')
    sleep(1)
    try:
        (ibrowser.click_by_xpath
         ('/html/body/div[3]/div[2]/div/mat-dialog-container/websbor-simple-dialog/websbor-base-dialog/div[2]/button'))
    except Exception:
        pass

    try:
        ibrowser.click_by_xpath('//input[@id="mat-input-3"]')
        ibrowser.click_by_xpath(f'//mat-option[@id="mat-option-{filial + 3}"]')
    except Exception:
        pass
    sleep(1)
    oktmo_get.append(ibrowser.driver.find_element(By.XPATH,
                                                  '//table[@class="table table-striped"]/tbody/tr[7]/td').text[:11])
    print('Пауза для росстата')
    sleep(10)

"""
ibrowser.quit()

# Изменить на путь куда файлы сохраняются по умолчанию
dir_path = 'C:/Users/Ярослав/Downloads'
filial = 0
check_remove = False

for count in range(len(inn_list)):
    pdf_path = find_pdf(dir_path, ogrn_get[count])
    okved_get.append(find_okved(pdf_path))
    if count != 0:
        if inn_list[count] == inn_list[count - 1]:
            filial += 1
        else:
            filial = 0
    if count != len(inn_list) - 1:
        if inn_list[count] == inn_list[count + 1]:
            check_remove = False
        else:
            check_remove = True
    if count == len(inn_list) - 1:
        check_remove = True
    kpp_get.append(find_kpp(pdf_path, filial, check_remove))
for i in range(len(inn_list)):
    print(f'ИНН: {inn_list[i]}, ogrn: {ogrn_get[i]} KPP: {kpp_get[i]} OKVED: {okved_get[i]}')

data_result = pandas.DataFrame({
    'ИНН': inn_list,
    'КПП исходный': kpp_isxod,
    'КПП получен': kpp_get,
    'OKVED исходный': okved_isxod,
    'OKVED получен': okved_isxod,
})
"""    'OKTMO исходный': oktmo_isxod,
    'OKTMO получен': oktmo_get,"""

data_result.to_excel('documents/result.xlsx')

finish_time = time.time()
print(finish_time - time_start)
