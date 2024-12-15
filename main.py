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


# Read CSV file skipping first 3 rows and using 4th row as header
df = pd.read_csv('tmp.csv', skiprows=3, encoding='utf-8-sig')

# Convert money columns to float
df['outcome'] = df['outcome'].apply(parse_money)
df['income'] = df['income'].apply(parse_money)

# Create summary by category
summary = pd.pivot_table(
    df,
    index='categoryName',
    values=['outcome', 'income'],
    aggfunc='sum',
    fill_value=0
)

# Sort by total money movement (outcome + income)
summary['total'] = summary['outcome'] + summary['income']
summary = summary.sort_values('total', ascending=False)

# Format numbers for better readability
summary = summary.round(2)

# Save results to CSV
summary.to_csv('summary.csv', encoding='utf-8-sig')

# Print results
print("\nCategory Summary (sorted by total money movement):")
print("=" * 80)
print(summary)
