import pandas as pd
import calendar

# Function to process data for a given month
def process_month(month):
    # Read Excel files into DataFrames
    month_file = f'/Users/shiggy/Desktop/{month.lower()}.xlsx'
    month_copart_file = f'/Users/shiggy/Desktop/{month.lower()}_copart.xlsx'

    df1 = pd.read_excel(month_file)
    df2 = pd.read_excel(month_copart_file)

    # Find VINs exclusive to the current month and its CoPart file
    not_in_copart = set(df1['VIN Number']).difference(set(df2['VIN Number']))
    copart_not_in_month = set(df2['VIN Number']).difference(set(df1['VIN Number']))

    # Create DataFrames for each category
    not_in_copart_df = df1[df1['VIN Number'].isin(not_in_copart)]
    copart_not_in_month_df = df2[df2['VIN Number'].isin(copart_not_in_month)]

    # Write the non-common VINs DataFrames to new Excel files
    not_in_copart_file = f'/Users/shiggy/Desktop/in_{month.lower()}_not_in_{month.lower()}_copart.xlsx'
    copart_not_in_month_file = f'/Users/shiggy/Desktop/in_{month.lower()}_copart_not_in_{month.lower()}.xlsx'

    not_in_copart_df.to_excel(not_in_copart_file, index=False)
    copart_not_in_month_df.to_excel(copart_not_in_month_file, index=False)

    print(f"Process for {month} completed successfully.")

# Iterate over the months
for month_num in range(1, 13):
    month_name = calendar.month_abbr[month_num]
    process_month(month_name)
