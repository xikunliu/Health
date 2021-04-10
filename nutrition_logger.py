# A Python script to update the nutrition_logger.xlsx with summaries
import pandas as pd # Load libraries
import numpy as np
import openpyxl

# Global variables
WORKBOOK_NAME = 'nutrition_log.xlsx'
DAILY_SUMMARY_SPREADSHEET_NAME = 'nutrition_summaries'

### Load workbooks and spreadsheets
nutrition_log = pd.ExcelFile(WORKBOOK_NAME)

### Load food_logger
food_logger = pd.read_excel(nutrition_log, 'food_logger', skiprows=[0,1])
food_logger = food_logger.iloc[:,0:4].dropna()

### Load food_nutrition
# Skip the title, blank row, and the egg scrambled/over easy
food_nutrition = pd.read_excel(nutrition_log, 'food_nutrition', skiprows=[0,1])
food_nutrition = food_nutrition[:-1] # Drop last row with the saved function

# Remove extra blank rows
food_nutrition.dropna(axis=0, how='all', inplace=True)

# Subset the notes on foods
food_nutrition_notes = food_nutrition.loc[:,'Notes']
food_nutrition = food_nutrition.drop(columns=['Notes'])

# Fill NA's with 0
food_nutrition.fillna(0, inplace=True)

### Load daily_values
# Skip the title, blank row, and the egg scrambled/over easy
daily_values = pd.read_excel(nutrition_log, 'daily_values', skiprows=[0,1])
daily_values = daily_values.iloc[:,0:2]

### Calculate nutrition_summaries
### Sum the nutrition values per day
unique_dates = food_logger.Date.unique()
daily_summary_list = []
for day in unique_dates:
    # Reference: https://stackoverflow.com/questions/10373660/converting-a-pandas-groupby-output-from-series-to-dataframe
    # For each day in food_logger, group the foods by Name and sum
    # the Total Servings
    daily_groupby = food_logger[food_logger.Date == day].loc[\
        :,['Name', 'Total Servings']].groupby(['Name'], as_index=False).sum()
    
    grouped_food = daily_groupby.iloc[1,:] # The unique food names in the daily_groupby
    
    # Reference: https://stackoverflow.com/questions/45576800/how-to-sort-dataframe-based-on-a-column-in-another-dataframe-in-pandas
    # Re-index the groupby so that it matches the order of foods in food_nutrition
    daily_groupby = daily_groupby.set_index('Name')
    daily_groupby = daily_groupby.reindex(index=food_nutrition['Name'])
    daily_groupby = daily_groupby.reset_index()
    daily_groupby.dropna(inplace=True) # Drop added foods that are not in daily_groupby
    
    # Reference: https://stackoverflow.com/questions/22542312/slice-pandas-dataframe-where-columns-value-exists-in-another-array
    # Query the corresponding foods from food_nutrition
    fn_query = food_nutrition.query("Name in @daily_groupby.Name")
    
    # Reference: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.multiply.html
    # Multiply each of the nutrition values by their food_logger.'Total Servings'
    # and sum column-wise
    daily_nutrition_sum = fn_query.iloc[:,1:].mul(daily_groupby.loc[:, 'Total Servings'], 0).sum(axis=0)
    daily_summary_list.append(daily_nutrition_sum)

### Summarize daily results
# Clean the RDA from daily_values
daily_values_rda = daily_values.loc[:,'Daily Value']
daily_values_rda.fillna(0, inplace=True)
daily_values_rda.index = daily_values.loc[:,'Nutrient'] # Allows for series division

daily_nutrient_percentage_list = []
for daily_summary in daily_summary_list:
    # Calculate percentages, convert Inf and NaN's to -1
    daily_nutrient_percentages = round((daily_summary[1:] / daily_values_rda), 4)
    daily_nutrient_percentages.fillna(-1e-2, inplace=True)
    daily_nutrient_percentages = daily_nutrient_percentages.replace(np.inf, -1e-2)
    daily_nutrient_percentage_list.append(daily_nutrient_percentages)

# Convert to dataframe and add Date column
daily_nutrition_summary = pd.DataFrame(daily_nutrient_percentage_list)
daily_nutrition_summary = pd.concat([pd.DataFrame({'Date': unique_dates}), daily_nutrition_summary], axis=1)

# Open workbook and spreadsheet
workbook = openpyxl.load_workbook(filename=WORKBOOK_NAME)
nutrition_summaries = workbook[DAILY_SUMMARY_SPREADSHEET_NAME]

# Reference: https://openpyxl.readthedocs.io/en/stable/pandas.html
# Reference: https://stackoverflow.com/questions/36657288/copy-pandas-dataframe-to-excel-using-openpyxl
# Save the daily summary to the nutrition_summaries spreadsheet from A4
from openpyxl.utils.dataframe import dataframe_to_rows
rows = dataframe_to_rows(daily_nutrition_summary, index=False, header=False)

for r_idx, row in enumerate(rows, 1):
    for c_idx, value in enumerate(row, 1):
        nutrition_summaries.cell(row=r_idx + 3, column=c_idx, value=value)

workbook.save(filename=WORKBOOK_NAME) # Save results