import pandas as pd
import os

ymaps_file = 'phone-call-list-3421500962632356867-2024-04-10T14_29_13.xlsx'
yc_file = 'Сетевые-клиенты.xlsx'
kudrovo = 'Клиенты-Кудрово.xlsx'
ohta = 'Клиенты-Охта.xlsx'


# удаляю дубли в таблице Я.Карт
ymaps_file_new = 'Новая таблица из Я Карт.xlsx'

df = pd.read_excel(ymaps_file)
df = df.drop(df.index[range(6888, len(df))]) # закомментируй, если не надо удалять лишние строки
df = df.drop_duplicates(subset=['Номер звонившего'])
df.rename(columns={'Номер звонившего': 'Телефон'}, inplace=True) # заменяю заголовок на Телефон
df.to_excel(ymaps_file_new, index=False, engine='openpyxl')


# объединяю 2 базы (из YC и Я.Карт) в 1 в txt формате
total_base = 'total_base_YC_YM.txt'
files = [yc_file, ymaps_file_new]

with open(total_base, 'w') as file_with_numbers:
    for file in files:
        df = pd.read_excel(file, engine='openpyxl')
        df['Телефон'] = df['Телефон'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))
        df['Телефон'] = df['Телефон'].apply(lambda x: '7' + x if not x.startswith('7') else x)

        for number in df['Телефон']:
            file_with_numbers.write(number + '\n')
print("Номера телефонов из всех файлов успешно сохранены в файл total_base_YC_YM.txt.")

os.remove(yc_file)
os.remove(ymaps_file_new)


# удаляю дубликаты из общей базы
with open(total_base, 'r') as file:
    numbers = [line.strip() for line in file]
counts = {}
for number in numbers:
    counts[number] = counts.get(number, 0) + 1
dublicates = {number: count for number, count in counts.items() if count > 1}
if dublicates:
    with open('base.txt', 'w') as clean_base:
        clean_base.write(f"Общее количество дубликатов: {int(sum(dublicates.values()) / 2)}\n")
        for number, count in dublicates.items():
            clean_base.write(f"{number}\n")
else:
    print("Дубликаты номеров телефонов не найдены.")

os.remove(total_base)


# фильтрую таблицу из YC по данным из отфильтрованной базы по Кудрово и Охте
with open('base.txt', 'r') as file:
    phone_number = set(line.strip() for line in file)
df = pd.read_excel(kudrovo, engine='openpyxl')
filtered_df = df[df['Телефон'].astype(str).apply(lambda x: x in phone_number)]
filtered_df.to_excel('Отфильтрованные Клиенты Кудрово.xlsx', index=False, engine='openpyxl')

with open('base.txt', 'r') as file:
    phone_number = set(line.strip() for line in file)
df = pd.read_excel(ohta, engine='openpyxl')
filtered_df = df[df['Телефон'].astype(str).apply(lambda x: x in phone_number)]
filtered_df.to_excel('Отфильтрованные Клиенты Охта.xlsx', index=False, engine='openpyxl')

os.remove(kudrovo)
os.remove(ohta)