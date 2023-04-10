import pandas as pd

input_file = pd.read_excel('task 2.xlsx')
df = pd.DataFrame(input_file)
df.drop(df.columns[[1, 2]], axis=1, inplace=True)

df.rename(columns={df.columns[0]: 'Філія'}, inplace=True)
column_department = df.columns[0]
df.rename(columns={df.columns[1]: 'Дата'}, inplace=True)
column_date = df.columns[1]
column_year = 'Рік'
column_week = 'Номер тижня'
column_files_per_week = 'Кількість файлів за тиждень'
column_files_per_day = 'Кількість файлів за день'

df[column_date] = pd.to_datetime(df[column_date], format='%Y-%m-%d')
df[column_year] = pd.DatetimeIndex(df[column_date]).year
df[column_week] = df[column_date].dt.isocalendar().week
df[column_date] = pd.to_datetime(df[column_date]).dt.date

for i in df.index:
    path_substring = df.at[i, column_department][3:]
    first_slash = path_substring.find('\\')
    path_substring = path_substring[:first_slash]
    df.at[i, column_department] = path_substring

df[column_files_per_week] = 0
df[column_files_per_week] = df.groupby([column_department, column_year, column_week])[column_date].transform('count')
df.drop_duplicates()

df[column_files_per_day] = 0
columns_result = [column_department, column_year, column_week, column_files_per_week, column_date]
df = df.groupby(columns_result)[column_files_per_day].agg('count')

df.to_excel('result.xlsx')
