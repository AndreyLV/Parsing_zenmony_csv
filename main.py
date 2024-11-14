import pandas as pd

# Загрузка CSV файла
df = pd.read_csv("tmp.csv")  # Замените на путь к вашему файлу

# Создание сводной таблицы
pivot_table = pd.pivot_table(
    df,
    values="outcome",  # Столбец для агрегирования (например, суммы)
    index="categoryName",  # Столбец для строк сводной таблицы
    columns="categoryName",  # Столбец для заголовков столбцов
    aggfunc="sum"  # Агрегирующая функция, например, сумма или среднее
)

print(pivot_table)
