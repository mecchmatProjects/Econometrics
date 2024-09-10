import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt

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

OLD_BORDER = 1990

old_tracks = len(column_data[column_data < OLD_BORDER])
print("Olds", old_tracks, old_tracks / len(column_data))
# Plot the histogramplt.figure(figsize=(10, 6))
plt.hist(column_data, bins=range(int(column_data.min()), int(column_data.max()) + 1, 5), color='skyblue', edgecolor='black')

# Mark the median and mode on the histogram
plt.axvline(median, color='red', linestyle='dashed', linewidth=1, label=f'Median: {median}')
plt.axvline(mode, color='green', linestyle='dashed', linewidth=1, label=f'Mode: {mode}')
plt.axvline(average, color='cyan', linestyle='dashed', linewidth=1, label=f'Avg: {int(average)}')


# Set title and labels
plt.title('Histogram of Years in Column Data')
plt.xlabel('Year')
plt.ylabel('Frequency')
plt.legend()
plt.show()