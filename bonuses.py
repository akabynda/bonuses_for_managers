import pandas as pd

# Загрузка данных из файла Excel
file_path = 'data.xlsx'
data = pd.read_excel(file_path)

# Очистка данных: удаление ненужных столбцов и преобразование дат
data = data.drop(columns=['Unnamed: 5'])
data['receiving_date'] = pd.to_datetime(data['receiving_date'], errors='coerce')
data['year'] = data['receiving_date'].dt.year
data['month'] = data['receiving_date'].dt.month

# Фильтрация сделок до 01.07.2021
eligible_deals = data[(data['year'] < 2021) | ((data['year'] == 2021) & (data['month'] < 7))].copy()

# Определение функции для расчета бонусов
def calculate_bonus(row):
    if row['document'] == 'оригинал':
        if row['new/current'] == 'новая' and row['status'] == 'ОПЛАЧЕНО':
            return row['sum'] * 0.07
        elif row['new/current'] == 'текущая' and row['status'] != 'ПРОСРОЧЕНО':
            return row['sum'] * 0.05 if row['sum'] > 10000 else row['sum'] * 0.03
    return 0

# Применение функции к каждой строке
eligible_deals['bonus'] = eligible_deals.apply(calculate_bonus, axis=1)

# Фильтрация сделок, оригиналы для которых пришли позже заключения сделки
delayed_contracts = eligible_deals[eligible_deals['receiving_date'] > pd.to_datetime(eligible_deals[['year', 'month']].assign(day=1))]

# Группировка по менеджерам и суммирование бонусов
bonuses_by_manager = delayed_contracts.groupby('sale')['bonus'].sum().reset_index()

# Переименование столбцов для удобства
bonuses_by_manager.columns = ['Менеджер', 'Остаток бонусов на 01.07.2021']

# Вывод результатов с красивой печатью
print(bonuses_by_manager.to_string(index=False))
