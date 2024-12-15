import pandas as pd
from scipy import stats
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import rankdata, kendalltau, chi2_contingency, mannwhitneyu

# Load the Excel file
file_path = 'Age of the trucks.xlsx'
df = pd.read_excel(file_path)

# Assume the numbers are in the first column
column_data = df.iloc[:, 0].dropna()
print(column_data, column_data.shape)

# Calculate the median
median = column_data.median()
average = column_data.mean()

# Calculate the mode
mode = stats.mode(column_data)[0]

print(f"Median: {median}")
print(f"Mode: {mode}")
print(f"Average: {average}")
print(column_data.min(), column_data.max())

data = column_data.to_numpy()

OLD_BORDER = 1990

old_tracks = len(column_data[column_data < OLD_BORDER])
print("Olds", old_tracks, old_tracks / len(column_data))
# # Plot the histogramplt.figure(figsize=(10, 6))
# plt.hist(column_data, bins=range(int(column_data.min()), int(column_data.max()) + 1, 5), color='skyblue', edgecolor='black')
#
# # Mark the median and mode on the histogram
# plt.axvline(median, color='red', linestyle='dashed', linewidth=1, label=f'Median: {median}')
# plt.axvline(mode, color='green', linestyle='dashed', linewidth=1, label=f'Mode: {mode}')
# plt.axvline(average, color='cyan', linestyle='dashed', linewidth=1, label=f'Avg: {int(average)}')
#
#
# # Set title and labels
# plt.title('Histogram of Romania tracks ages')
# plt.xlabel('Year')
# plt.ylabel('Number of tracks')
# plt.legend()
# plt.show()


file_path = "Data_questionaire_finance_tracksyears.xlsx"

df = pd.read_excel(file_path,sheet_name="TracsYears")

column_data2 = df.iloc[1:84, 1:26+17]

print(column_data2, column_data2.shape)
data2 = column_data2.to_numpy().ravel()
data2 = data2[~np.isnan(data2)]
print(data2, data2.shape)

# Perform the Mann - Whitney U-Test

stat_mu, p_value_mu = mannwhitneyu(data, data2)

# Output the results
print(f"Mann-Whitney U Test statistic: {stat_mu}")
print(f"P-value: {p_value_mu}")



