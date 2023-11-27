from python_pptx_text_replacer import TextReplacer
import csv
import os
import time


file_example = 'Шаблон.pptx'
directory = 'SLIDES'

if not os.path.isdir(directory):
    os.mkdir(directory)

with open('CSV.csv', 'r', newline='', encoding='windows-1251') as file:
    reader = csv.reader(file, delimiter=';')
    next(reader)
    num_slide = 0

    for row in reader:
        replacer = TextReplacer(file_example, slides='', tables=True, charts=True, textframes=True)
        num_slide += 1
        replacer.replace_text([('Name', row[1]), ('Cena1', row[11]), ('Cena2', row[9]), ('Ahour', row[2]), ('Amper', row[3]), ('Polar', row[4]), ('Razmer', row[5]), ('Garant', row[6]), ('Strana', row[7]), ('Zavod', row[8]), ('Cena3', row[10])])

        new_file = f'{directory}/{num_slide}.pptx'
        replacer.write_presentation_to_file(new_file)

print('''
--------------------------------------------------------------------------------
Программа "CardPricePPTX" предназначена для пакетного создания ценников товаров.
По шаблону презентации PowerPoint (файл Шаблон.pptx) методом замены значимой информации в директории SLIDES создаются файлы, в каждом из которых имеется 1 слайд с ценником.
Слайды затем можно штатными средствами PowerPoint собрать в одну презентацию для дальнейшей печати.
Данные для замены должны находиться в файле CSV.csv. Порядок столбцов не менять. Он важен для корректной замены.''')
time.sleep(1)
print(f'''
----------------------------------------------------------------
Работа программы завершена. Программа закроется через 20 секунд.

В папке SLIDES созданы пронумерованные файлы ({num_slide} шт.), в каждом из которых имеется 1 слайд с ценником.
Слайды затем можно штатными средствами PowerPoint собрать в одну презентацию для дальнейшей печати.

ENJOY!
------------------------------------------------------------------------------
Замечания по работе программы, предложения, пожелания, благодарности и донаты:
\tTelegram: @iuriigav
\tEmail: yubarssoft@yandex.ru''')

time.sleep(20)
