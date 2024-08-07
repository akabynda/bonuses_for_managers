import pandas as pd
import matplotlib.pyplot as plt

# Загрузка данных из файла Excel
file_path = 'data.xlsx'
data = pd.read_excel(file_path)

# Очистка данных: удаление ненужных столбцов и преобразование дат
data = data.drop(columns=['Unnamed: 5'])
data['receiving_date'] = pd.to_datetime(data['receiving_date'], errors='coerce')
data['year'] = data['receiving_date'].dt.year
data['month'] = data['receiving_date'].dt.month

# 1) Вычисление общей выручки за июль 2021 по тем сделкам, приход денежных средств которых не просрочен
july_2021_data = data[(data['year'] == 2021) & (data['month'] == 7)]
non_overdue_july_2021 = july_2021_data[july_2021_data['status'] != 'ПРОСРОЧЕНО']
total_revenue_july_2021 = non_overdue_july_2021['sum'].sum()

# 2) Как изменялась выручка компании за рассматриваемый период? Проиллюстрировать графиком
monthly_revenue = data.groupby(['year', 'month'])['sum'].sum().reset_index()
monthly_revenue['date'] = pd.to_datetime(monthly_revenue[['year', 'month']].assign(day=1))

plt.figure(figsize=(12, 6))
plt.plot(monthly_revenue['date'], monthly_revenue['sum'], marker='o', linestyle='-')
plt.title('Изменение выручки компании по месяцам')
plt.xlabel('Дата')
plt.ylabel('Общая выручка')
plt.grid(True)
plt.savefig('monthly_revenue.png')  # Сохранение графика в файл
# plt.show()

# 3) Кто из менеджеров привлек для компании больше всего денежных средств в сентябре 2021?
september_2021_data = data[(data['year'] == 2021) & (data['month'] == 9)]
september_2021_revenue_by_manager = september_2021_data.groupby('sale')['sum'].sum().reset_index()
top_manager_september_2021 = september_2021_revenue_by_manager.loc[september_2021_revenue_by_manager['sum'].idxmax()]

# 4) Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
october_2021_data = data[(data['year'] == 2021) & (data['month'] == 10)]
deal_types_october_2021 = october_2021_data['new/current'].value_counts()

# 5) Сколько оригиналов договора по майским сделкам было получено в июне 2021?
may_2021_data = data[(data['year'] == 2021) & (data['month'] == 5)]
# Фильтрация майских сделок и получение оригиналов договоров в июне
contracts_received_june_2021 = may_2021_data[(may_2021_data['receiving_date'].dt.year == 2021) & (may_2021_data['receiving_date'].dt.month == 6)]
num_contracts_received_june_2021 = contracts_received_june_2021['document'].value_counts().get('оригинал', 0)

# Печать результатов
print("Результаты анализа данных:")
print("=" * 50)
print(f"Общая выручка за июль 2021 (не просроченные сделки): {total_revenue_july_2021:.2f}")
print(f"Менеджер, привлекший больше всего денежных средств в сентябре 2021: {top_manager_september_2021['sale']} - {top_manager_september_2021['sum']}")
print(f"Преобладающий тип сделок в октябре 2021: {deal_types_october_2021.idxmax()} - {deal_types_october_2021.max()}")
print(f"Количество оригиналов договора, полученных в июне 2021 по майским сделкам: {num_contracts_received_june_2021}")
print("=" * 50)
