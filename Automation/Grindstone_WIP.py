import glob
import numpy as np
import pandas as pd
from glob import glob
import os
files = ['Company 1 Profit & Loss November 2020.xlsx', 
          'Company 2 Profit & Loss November 2020.xlsx',
          'Company 3 Profit & Loss November 2020.xlsx',
          'Company 4 Profit & Loss November 2020.xlsx',
          'Company 5 Profit & Loss November 2020.xlsx',
          'Company 6 Profit & Loss November 2020.xlsx']


# path = r'E:\Python Projects\Grindstone\Variance Report Data - Output Example\Example Input Data'
path= r'E:\Python Projects\Grindstone\Variance Report Data - Output Example\Example Input Data'

              
all_files = glob.glob(os.path.join(path, "*.xlsx"))     # advisable to use os.path.join as this makes concatenation OS independent


# all_files =  glob.glob(os.path.join(path, "*.csv"))
# # filenames = glob.glob(os.path.join(path, "*.xlsx"))
all_files  = glob.glob(path + "/*.xlsx")


df_from_each_file = (pd.read_csv(f) for f in all_files)
dfs = [pd.read_excel(filename) for filename in all_files]


df=pd.read_excel(r'E:\Python Projects\Grindstone\Variance Report Data - Output Example\Example Input Data\Company_1.xlsx')
df.dropna(how='all', inplace=True)

filename = df['Profit & Loss'].iloc[0]
df_PL = df.iloc[1:-1]


index_Revenue = df_PL.loc[df_PL['Profit & Loss'] == 'Revenue'].index[0]
index_TotalRevenue = df_PL.loc[df_PL['Profit & Loss'] == 'Total Revenue'].index[0]
TotalRevenue = df_PL.iloc[index_Revenue: index_TotalRevenue-1]
TotalRevenue.set_index("Profit & Loss", inplace =True)
TotalRevenue = TotalRevenue.rename(columns=  {'Unnamed: 1': 'Oct 2020','Unnamed: 2': 'Nov 2020','Unnamed: 3': 'Variance ($)', 'Unnamed: 4': 'Variance (%)' }).fillna('')
TotalRevenue= TotalRevenue.append(TotalRevenue[['Oct 2020', 'Nov 2020','Variance ($)','Variance (%)']].sum().rename('Total Revenue')).fillna('')
total_revenue = TotalRevenue.loc[TotalRevenue.index == "Total Revenue"]
total_revenue


#cost of sales 
index_CostofSales = df_PL.loc[df_PL['Profit & Loss'] == 'Cost of Sales'].index[0]
index_TotalCostofSales = df_PL.loc[df_PL['Profit & Loss'] == 'Total Cost of Sales'].index[0] 
# COS = df_PL.iloc[index_CostofSales: index_TotalCostofSales, 1:]
COS = df_PL.iloc[index_CostofSales: index_TotalCostofSales]


index_Expenses = df_PL.loc[df_PL['Profit & Loss'] == 'Expenses'].index[0]
index_TotalExpenses = df_PL.loc[df_PL['Profit & Loss'] == 'Total Expenses'].index[0]
Expenses = df_PL.iloc[index_Expenses: index_TotalExpenses]
Expenses

index_OtherExpenses = df_PL.loc[df_PL['Profit & Loss'] == 'Other Expenses'].index[0]
index_TotalOtherExpenses = df_PL.loc[df_PL['Profit & Loss'] == 'EBITDA'].index[0]
OtherExpenses = df_PL.iloc[index_OtherExpenses: index_TotalOtherExpenses-1, 1:]
OtherExpenses

df_PL.columns = df_PL.iloc[0]
df_PL.set_index(np.nan,inplace=True)
#df_PL.fillna(0, inplace=True)


main_expenses = Expenses[(Expenses['Unnamed: 3'] <= 5000) & (Expenses['Unnamed: 3'] >= -5000)&(Expenses['Unnamed: 4'] <= 5) & (Expenses['Unnamed: 4'] >= -5)]. rename(columns=  {'Unnamed: 1': 'Oct 2020','Unnamed: 2': 'Nov 2020','Unnamed: 3': 'Variance ($)', 'Unnamed: 4': 'Variance (%)' })

index_TotalRevenue = df_PL.loc[df_PL['Profit & Loss'] == 'Total Revenue'].index[0]
index_NetIncome= df_PL.loc[df_PL['Profit & Loss'] == 'Net Income'].index[0]

main= Expenses[(Expenses['Unnamed: 3'] <= 5000) & (Expenses['Unnamed: 3'] >= -5000)&(Expenses['Unnamed: 4'] <= 5) & (Expenses['Unnamed: 4'] >= -5)]
#Expenses.append(pd.Series(new_row, index=df.columns, name='e'))
#a.to_excel(r'E:\Python Projects\Grindstone\Trial.xlsx', index=False)

a = Expenses.append(Expenses[['Oct 2020', 'Nov 2020','Variance ($)','Variance (%)']].sum().rename('Subtotal Main Expenses'))
a
          






