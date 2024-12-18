import pandas as pd


def parse_money(value):
    """Parse money values like '1.234,56' to float"""
    if pd.isna(value):
        return 0.0
    try:
        # Remove currency symbol if present and strip whitespace
        cleaned = str(value).strip().replace('RUB', '').strip()
        # Replace comma with dot and remove thousand separators
        cleaned = cleaned.replace(',', '.').replace(' ', '')
        return float(cleaned)
    except ValueError:
        return 0.0


def format_money(value):
    """Format float to Russian style money format"""
    return f"{value:,.2f}".replace(",", " ")


# Read CSV file skipping first 3 rows and using 4th row as header
df = pd.read_csv('tmp.csv', skiprows=3, encoding='utf-8-sig')

# Convert money columns to float
df['outcome'] = df['outcome'].apply(parse_money)
df['income'] = df['income'].apply(parse_money)

# Create summary by category
summary = pd.pivot_table(
    df,
    columns='categoryName',
    values=['outcome', 'income'],
    aggfunc='sum',
    fill_value=0
)

# Sort categories by total money movement
total_movement = summary.sum()
sorted_categories = total_movement.sort_values(ascending=False).index

# Reorder columns according to total movement
summary = summary[sorted_categories]

# Add total row
summary.loc['total'] = summary.loc['income'] + summary.loc['outcome']

# Format numbers
for col in summary.columns:
    summary[col] = summary[col].apply(format_money)

# Rename index for better readability
summary.index = ['Доходы', 'Расходы', 'Итого']

# Calculate totals for summary sheet
total_income = df['income'].sum()
total_outcome = df['outcome'].sum()
total_balance = total_income - total_outcome

# Create summary DataFrame
result_summary = pd.DataFrame({
    'Тип': ['Доходы', 'Расходы', 'Итого'],
    'Значения': [
        format_money(total_income),
        format_money(total_outcome),
        format_money(total_balance)
    ]
})

# Save results to Excel
with pd.ExcelWriter('financial_report.xlsx', engine='openpyxl') as writer:
    # Write summary sheet
    result_summary.to_excel(writer, sheet_name='Итог', index=False)

    # Write detailed sheets with categories in columns
    summary.to_excel(writer, sheet_name='Категории')

    # Adjust columns width
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (TypeError, AttributeError):
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

# Print results
print("\nФинансовый отчет:")
print("=" * 50)
print(result_summary.to_string(index=False))

print("\nДетальный отчет по категориям сохранен в файл 'financial_report.xlsx'")
print("Лист 'Категории' содержит:")
print("- Строка 'Расходы': сумма расходов по каждой категории")
print("- Строка 'Доходы': сумма доходов по каждой категории")
print("- Строка 'Итого': общая сумма движения денег по категории")
