import pandas as pd
import datetime
import numpy as np

# Set start date
d = datetime.date(2021,3,17)
StartDate = d.strftime("%Y%m%d")

# Read the first Excel file
df = pd.read_excel(r'C:\Users\lenad\OneDrive\Dokumenter\TPK4930 Masteroppgave\Datafiler\EXPORT_' + StartDate + '.xlsx')

# Create new, empty data frame
Usage = pd.DataFrame()

# Add one column for index and one column for usage
Usage['Index'] = df['Material'].astype(str) + df['Plant']
Usage['Usage ' + StartDate] = df['Total Goods Issue Quantity']

# Change row index
Usage = Usage.set_index('Index')

# Get the Usage-column from each Excel file and add it to the Usage data frame
i = 0
for i in range(5):
    d = d + datetime.timedelta(days=7)
    StartDate = d.strftime("%Y%m%d")
    df = pd.read_excel(r'C:\Users\lenad\OneDrive\Dokumenter\TPK4930 Masteroppgave\Datafiler\EXPORT_' + StartDate + '.xlsx')
    df['Index'] = df['Material'].astype(str) + df['Plant']
    df = df.set_index('Index')
    df_usage = df['Total Goods Issue Quantity']
    Usage = Usage.join(df_usage)
    Usage = Usage.rename(columns={'Total Goods Issue Quantity': 'Usage ' + StartDate})
    i = i + 1

Usage = Usage.groupby(Usage.index).sum()

# Replace NaN with 0 og print the data frame
Usage = Usage.replace(np.nan,0)

# Sort the data frame by ascending order
Usage['Material number and plant'] = Usage.index
# Usage = Usage.sort_values(by=['Material number and plant'])

# Save data frame to new Excel file named "Usage"
print(Usage)
Usage.to_excel(r'C:\Users\lenad\OneDrive\Dokumenter\TPK4930 Masteroppgave\Datafiler\Usage.xlsx', sheet_name='Usage', index=False)
