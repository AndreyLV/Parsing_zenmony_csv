import pandas as pd


def parse_money(value):
    """Parse money values like '1.234,56' to float"""
    if pd.isna(value):
        return 0.0
    try:
        # Remove currency symbol if present and strip whitespace
        cleaned = str(value).strip().replace("RUB", "").strip()
        # Replace comma with dot and remove thousand separators
        cleaned = cleaned.replace(",", ".").replace(" ", "")
        return float(cleaned)
    except ValueError:
        return 0.0


# Read CSV file skipping first 3 rows and using 4th row as header
df = pd.read_csv("tmp.csv", skiprows=3, encoding="utf-8-sig")

# Convert money columns to float
df["outcome"] = df["outcome"].apply(parse_money)
df["income"] = df["income"].apply(parse_money)

# remove #category
df["categoryName"] = df["categoryName"].str.replace(r", #.*", "", regex=True)

# rename incomin "Другое" to "Другое +"
df.loc[df["income"] != 0.00, "categoryName"] = df.loc[
    df["income"] != 0.00, "categoryName"
].str.replace(r"Другое", "Другое+", regex=True)

# rename columns
df["categoryName"] = df["categoryName"].str.replace(
    r"Дом, квартира.*", "Дом, квартира", regex=True
)
df["categoryName"] = df["categoryName"].str.replace(
    r"Correction", "Другое+", regex=True
)
df["categoryName"] = df["categoryName"].str.replace(
    r"Инвестиции", "Другое+", regex=True
)
df["categoryName"] = df["categoryName"].str.replace(r"Кэшбэк", "Другое+", regex=True)


# Create summary by category
summary = pd.pivot_table(
    df,
    columns="categoryName",
    values=["outcome", "income"],
    aggfunc="sum",
    fill_value=0,
)


# Sort categories by total money movement
total_movement = summary.sum()
sorted_categories = total_movement.sort_values(ascending=False).index

# Reorder columns according to total movement
# summary = summary[sorted_categories]

desired_order = [
    "Зарплата / Работа",
    "Зарплата / Доп. работа",
    "Другое+",
    "Проезд",
    "Еда на заказ",
    "Здоровье",
    "Дом, квартира",
    "Интернет",
    "Телефон",
    "Хоз. Товары",
    "Продукты",
    "Отдых",
    "Спорт/здоровье",
    "Подписки",
    "Подарки",
    "Другое",
    "Авто",
    "Отпуск",
    "Крупняк",
]

for col in desired_order:
    if col not in summary.columns:
        summary[col] = 0

summary = summary[desired_order]
# summary = summary[sorted_categories]

# Add total row
summary.loc["total"] = summary.loc["income"] + summary.loc["outcome"]

# Rename index for better readability
summary.index = ["Доходы", "Расходы", "Итого"]

# Save results to Excel
with pd.ExcelWriter("financial_report.xlsx", engine="openpyxl") as writer:
    # Write detailed sheets with categories in columns
    summary.to_excel(writer, sheet_name="Категории")

    # Adjust columns width
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for i in range(3, len(desired_order) * 2, 2):
            worksheet.insert_cols(idx=i)
        worksheet.insert_cols(idx=7)
        worksheet.insert_cols(idx=7)
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (TypeError, AttributeError):
                    pass
            adjusted_width = max_length + 2
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
