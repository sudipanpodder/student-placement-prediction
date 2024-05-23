import pandas as pd

file_path = 'year-of-graduation-calculation.xlsx'
data = pd.read_excel(file_path, sheet_name='Data')

data['Created'] = pd.to_datetime(data['Created'], infer_datetime_format=True)

data['Academic_Year'] = pd.to_numeric(data['Academic_Year'], errors='coerce').fillna(0).astype(int)

def calculate_expected_graduation_date(created_date, academic_year):
    year_created = created_date.year
    month_created = created_date.month
    expected_graduation_year = year_created + 4 - academic_year
    if month_created >= 7:
        expected_graduation_year += 1
    return expected_graduation_year

data['Calculated Expected Graduation Date'] = data.apply(
    lambda row: calculate_expected_graduation_date(row['Created'], row['Academic_Year']),
    axis=1
)

output_file_path = 'calculated_graduation_dates.xlsx'
data.to_excel(output_file_path, index=False)

print(f"Processed data saved to {output_file_path}")
