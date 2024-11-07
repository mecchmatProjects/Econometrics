import os
import time
#import openpyxl
import numpy as np
import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable
from scipy.stats import pearsonr, ttest_1samp

# FIRMS_DATA = "Chestionar_Viktor.xlsx"
FIRMS_DATA = "Chestionar_Transforemed.xlsx"
FIRMS_DATA2 = "Firms"

firm_data ={
    "name":"",
    "field":"",
    "employers":0,
    "progress":True,
    "perform":0
}

SHEET1 = 'Sheet1'

firms = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="A", dtype=str).dropna(how='all')
print(firms.values[1:],firms.index)
input()

avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
print(avg_stuff.values[:10])

stuff = []
stuffs = []
for i,peoples in enumerate(avg_stuff.values):
    print(peoples)
    if not pd.isnull(peoples):
        print("P", i, peoples, int(peoples))
        stuff.append(int(peoples[0]))
    elif i+1 in firms.index:
        print("ok")
        stuffs.append(stuff)
        stuff = []
    # input()
stuffs.append(stuff)
print(stuffs[:5])

L =[]
for i in range(len(firms)):
    print(i,firms.values[i],stuffs[i])
    L.append((str(firms.values[i][0]),stuffs[i][0]))

print(L[:10])



# Load the Excel file

df = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="O:U").dropna(how='all')
print(df[:10])

# Estimate Gross Profit (this is just a sample formula, you may adjust it based on your context)
df['Gross Profit'] = df['Turnover.1'] - (df['Liailities.1'] + df['Fixed assets.1'])

# Normalize Gross Profit to a 1 to 5 scale
min_gross_profit = df['Gross Profit'].min()
max_gross_profit = 0 # df['Gross Profit'].max()
print(min_gross_profit, max_gross_profit)
df['Gross Profit (1-5)'] = 1 + 4 * (df['Gross Profit'] - min_gross_profit) / (max_gross_profit - min_gross_profit)
print("Profit", df['Gross Profit (1-5)'])

# Preview the resulting DataFrame
print("Res" + df[['Turnover', 'Profit Net', 'Gross Profit', 'Gross Profit (1-5)']])

input()






avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
print(avg_stuff.values[:10])
profit = []
p = []
for i,pr in enumerate(avg_stuff.values):
    print(peoples)
    if not pd.isnull(peoples):
        print("P", i, peoples, int(peoples))
        stuff.append(int(peoples[0]))
    elif i+1 in firms.index:
        print("ok")
        stuffs.append(stuff)
        stuff = []
    # input()
stuffs.append(stuff)
print(stuffs[:5])

L =[]
for i in range(len(firms)):
    print(i,firms.values[i],stuffs[i])
    L.append((str(firms.values[i][0]),stuffs[i][0]))

print(L[:10])

TRANS = "Chestionar_Transforemed.xlsx"

years = pd.read_excel(TRANS, sheet_name="FirmAnalysis", usecols="X:AE", dtype=str).dropna(how='all')
print(years.values[:10])

years.to_excel("output.xlsx",
             sheet_name='Sheet_name_1')

input()

#
# avg_stuff = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, usecols="K", dtype=str)
# print(avg_stuff.values[:10])
# profit = []
# p = []
# for i,pr in enumerate(avg_stuff.values):
#     print(peoples)
#     if not pd.isnull(peoples):
#         print("P", i, peoples, int(peoples))
#         stuff.append(int(peoples[0]))
#     elif i+1 in firms.index:
#         print("ok")
#         stuffs.append(stuff)
#         stuff = []
#     # input()
# stuffs.append(stuff)
# print(stuffs[:5])

# Read the Excel file
# df = pd.read_excel(file_path)

df = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, dtype=object)


# Find the indices of the empty rows which separate the firms
empty_indices = df[df.isnull().all(axis=1)].index
print(empty_indices)

# Add an index at the end to handle the last segment
empty_indices = firms.index[1:] #  empty_indices.append(pd.Index([len(df)]))


print(empty_indices.values)

# Initialize a list to store results
firm_results = []

# Initialize start index
start_idx = 1

# Iterate over the indices to separate the data for each firm
for end_idx in empty_indices:
    # Slice the DataFrame to get the firm's data
    firm_data = df.iloc[start_idx+1:end_idx-1].dropna(how='all')
    firm_data = firm_data.replace(u'\xa0', u'', regex=True).astype(float)
    print(firm_data[:5].values)


    if not firm_data.empty:
        # Calculate the required metrics
        print(firm_data['Turnover'])

        t1 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[-1].split()))
        t2 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[0].split()))
        print("T", t1, " and", t2, float(t1))
        turnover_growth = (t1 / t2) # ** (1 / len(firm_data)) - 1) * 100
        print("grows", turnover_growth)

        print(firm_data['Profit Net'])
        print(firm_data['Profit Net'].values)
        print(firm_data['Turnover'].values)
        avg_profit_margin = (firm_data['Profit Net'].values / firm_data['Turnover'].values).mean() * 100
        print(avg_profit_margin)
        liabilities_to_assets = (
                    firm_data['Liailities'] / (firm_data['Fixed assets'] + firm_data['Circulant Assets'])).mean()
        print(liabilities_to_assets)
        fixed_assets_growth = (firm_data['Fixed assets'].iloc[-1] / firm_data['Fixed assets'].iloc[0])  #** ( 1 / (len(firm_data)) - 1) * 100


        current_ratio = (firm_data['Circulant Assets'] / firm_data['Liailities']).mean()
        capital_reserves_growth = ((firm_data['Capitals and reserves'].iloc[-1] /
                                    firm_data['Capitals and reserves'].iloc[0])) # ** (1 / (len(firm_data) - 1)) - 1) * 100

        # Determine scores for each metric
        turnover_score = min(max((turnover_growth // 5) + 1, 1), 5)
        profit_margin_score = min(max((avg_profit_margin // 5) + 1, 1), 5)
        liabilities_score = min(max(((1 - liabilities_to_assets) * 10), 1), 5)
        fixed_assets_score = min(max((fixed_assets_growth // 2) + 1, 1), 5)
        current_ratio_score = min(max((current_ratio * 2.5), 1), 5)
        capital_reserves_score = min(max((capital_reserves_growth // 3) + 1, 1), 5)

        # Analyze employees trend (here assuming a simple average trend as positive growth)
        employee_trend = (firm_data['The average number of employees'].iloc[-1] -
                          firm_data['The average number of employees'].iloc[0]) / len(firm_data)
        employee_score = 5 if employee_trend > 0 else 3 if employee_trend == 0 else 1

        # Calculate final score as average of all scores
        final_score = round((turnover_score + profit_margin_score + liabilities_score + fixed_assets_score + current_ratio_score + capital_reserves_score + employee_score) / 7,0)

        # Store the firm's final score
        firm_results.append(final_score)

    # Update start index for next firm
    start_idx = end_idx + 1

print(firm_results)

# Read the Excel file
df = pd.read_excel(FIRMS_DATA, sheet_name=SHEET1, dtype=object)


# Find the indices of the empty rows which separate the firms
empty_indices = df[df.isnull().all(axis=1)].index
print(empty_indices)

# Add an index at the end to handle the last segment
empty_indices = firms.index[1:] #  empty_indices.append(pd.Index([len(df)]))


print(empty_indices.values)

# Initialize a list to store results
firm_results = []

# Initialize start index
start_idx = 1

# Iterate over the indices to separate the data for each firm
for end_idx in empty_indices:
    # Slice the DataFrame to get the firm's data
    firm_data = df.iloc[start_idx+1:end_idx-1].dropna(how='all')
    firm_data = firm_data.replace(u'\xa0', u'', regex=True).astype(float)
    print(firm_data[:5].values)


    if not firm_data.empty:
        # Calculate the required metrics
        print(firm_data['Turnover'])

        t1 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[-1].split()))
        t2 = firm_data['Turnover'].iloc[-1] # float("".join(firm_data['Turnover'].iloc[0].split()))
        print("T", t1, " and", t2, float(t1))
        turnover_growth = (t1 / t2) # ** (1 / len(firm_data)) - 1) * 100
        print("grows", turnover_growth)

        print(firm_data['Profit Net'])
        print(firm_data['Profit Net'].values)
        print(firm_data['Turnover'].values)
        avg_profit_margin = (firm_data['Profit Net'].values / firm_data['Turnover'].values).mean() * 100
        print(avg_profit_margin)
        liabilities_to_assets = (
                    firm_data['Liailities'] / (firm_data['Fixed assets'] + firm_data['Circulant Assets'])).mean()
        print(liabilities_to_assets)
        fixed_assets_growth = (firm_data['Fixed assets'].iloc[-1] / firm_data['Fixed assets'].iloc[0])  #** ( 1 / (len(firm_data)) - 1) * 100


        current_ratio = (firm_data['Circulant Assets'] / firm_data['Liailities']).mean()
        capital_reserves_growth = ((firm_data['Capitals and reserves'].iloc[-1] /
                                    firm_data['Capitals and reserves'].iloc[0])) # ** (1 / (len(firm_data) - 1)) - 1) * 100

        # Determine scores for each metric
        turnover_score = min(max((turnover_growth // 5) + 1, 1), 5)
        profit_margin_score = min(max((avg_profit_margin // 5) + 1, 1), 5)
        liabilities_score = min(max(((1 - liabilities_to_assets) * 10), 1), 5)
        fixed_assets_score = min(max((fixed_assets_growth // 2) + 1, 1), 5)
        current_ratio_score = min(max((current_ratio * 2.5), 1), 5)
        capital_reserves_score = min(max((capital_reserves_growth // 3) + 1, 1), 5)

        # Analyze employees trend (here assuming a simple average trend as positive growth)
        employee_trend = (firm_data['The average number of employees'].iloc[-1] -
                          firm_data['The average number of employees'].iloc[0]) / len(firm_data)
        employee_score = 5 if employee_trend > 0 else 3 if employee_trend == 0 else 1

        # Calculate final score as average of all scores
        final_score = round((turnover_score + profit_margin_score + liabilities_score + fixed_assets_score + current_ratio_score + capital_reserves_score + employee_score) / 7,0)

        # Store the firm's final score
        firm_results.append(final_score)

    # Update start index for next firm
    start_idx = end_idx + 1

print(firm_results)

