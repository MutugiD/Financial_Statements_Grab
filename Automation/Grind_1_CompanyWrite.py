import pandas as pd 
import numpy as np 
import glob





def TotalRevenue_check(df): 
    col = df['Profit & Loss'].isin(['Total Revenue'])
    signal = []
    if col.any(): 
        signal = 1
    else: 
        signal = 0
      
def Revenue_line (df):
    signal = TotalRevenue_check(df)
    if signal == 1:
            index_Revenue = df.loc[df['Profit & Loss'] == 'Revenue'].index[0]
            index_TotalRevenue = df.loc[df['Profit & Loss'] == 'Total Revenue'].index[0]
            TotalRevenue = df.iloc[index_Revenue: index_TotalRevenue]
            TotalRevenue.set_index("Profit & Loss", inplace =True)
            TotalRevenue = TotalRevenue.rename(columns=  {'Unnamed: 1': df.iloc[0]['Unnamed: 1'],'Unnamed: 2': df.iloc[0]['Unnamed: 2'],'Unnamed: 3': df.iloc[0]['Unnamed: 3'],'Unnamed: 4': df.iloc[0]['Unnamed: 4'] }).fillna('')
            TotalRevenue = TotalRevenue.rename(columns=  {'Unnamed: 1': 'Oct 2020','Unnamed: 2': 'Nov 2020','Unnamed: 3': 'Variance ($)', 'Unnamed: 4': 'Variance (%)' }).fillna('')
            total_revenue = TotalRevenue.loc[TotalRevenue.index == "Total Revenue"]
            
    else:
        total_revenue = df.loc[df['Profit & Loss'] == 'Gross Profit Before Depreciation'].rename(columns=  {'Unnamed: 1': 'Oct 2020','Unnamed: 2': 'Nov 2020','Unnamed: 3': 'Variance ($)', 'Unnamed: 4': 'Variance (%)' }).fillna('')
        total_revenue.set_index("Profit & Loss", inplace =True)
        total_revenue= total_revenue.rename(index={'Gross Profit Before Depreciation': 'Total Revenue'})
    return total_revenue 

def Cost_of_Sales(df):
    index_CostofSales = df.loc[df['Profit & Loss'] == 'Cost of Sales'].index[0]
    index_TotalCostofSales = df.loc[df['Profit & Loss'] == 'Total Cost of Sales'].index[0] 
    cost_of_sales = df.iloc[index_CostofSales: index_TotalCostofSales-1]
    return cost_of_sales

def Main_Expenses(df):
    index_Expenses = df.loc[df['Profit & Loss'] == 'Expenses'].index[0]
    index_TotalExpenses = df.loc[df['Profit & Loss'] == 'Total Expenses'].index[0]
    Expenses = df.iloc[index_Expenses: index_TotalExpenses-1]
    
    expenses = Expenses[(Expenses['Unnamed: 3'] <= 5000) & (Expenses['Unnamed: 3'] >= -5000)&(Expenses['Unnamed: 4'] <= 5) & (Expenses['Unnamed: 4'] >= -5)].rename(columns= {'Profit & Loss': 'Ledger', 'Unnamed: 1': df.iloc[0]['Unnamed: 1'],'Unnamed: 2': df.iloc[0]['Unnamed: 2'],'Unnamed: 3': df.iloc[0]['Unnamed: 3'],'Unnamed: 4': df.iloc[0]['Unnamed: 4']}).fillna('')

    expenses.set_index('Ledger', inplace =True)
    main_expenses= expenses.append(expenses[[df.iloc[0]['Unnamed: 1'], df.iloc[0]['Unnamed: 2'], df.iloc[0]['Unnamed: 3'],df.iloc[0]['Unnamed: 4']]].sum().rename('Subtotal Main Expenses'))
    
    return main_expenses

def Other_Expenses(df):
    index_OtherExpenses = df.loc[df['Profit & Loss'] == 'Other Expenses'].index[0]
    index_TotalOtherExpenses = df.loc[df['Profit & Loss'] == 'EBITDA'].index[0]
    OtherExpenses = df.iloc[index_OtherExpenses: index_TotalOtherExpenses-1, 1:]
    other_expenses= OtherExpenses[(OtherExpenses['Unnamed: 3'] <= 5000) & (OtherExpenses['Unnamed: 3'] >= -5000)&(OtherExpenses['Unnamed: 4'] <= 5) & (OtherExpenses['Unnamed: 4'] >= -5)].rename(columns={'Profit & Loss': 'Ledger', 'Unnamed: 1': df.iloc[0]['Unnamed: 1'],'Unnamed: 2': df.iloc[0]['Unnamed: 2'],'Unnamed: 3': df.iloc[0]['Unnamed: 3'],'Unnamed: 4': df.iloc[0]['Unnamed: 4']}).fillna('')
    other_expenses = other_expenses.append(other_expenses[[df.iloc[0]['Unnamed: 1'], df.iloc[0]['Unnamed: 2'], df.iloc[0]['Unnamed: 3'],df.iloc[0]['Unnamed: 4']]].sum().rename('All Other Expenses'))
    other_expenses= other_expenses[other_expenses.index == 'All Other Expenses']
    return other_expenses


def Net_Income(df): 
    NetIncome = df.loc[df['Profit & Loss'] == 'Net Income']
    NetIncome.set_index("Profit & Loss", inplace =True)
    net_income = NetIncome.rename(columns=  {'Unnamed: 1': df.iloc[0]['Unnamed: 1'],'Unnamed: 2': df.iloc[0]['Unnamed: 2'],'Unnamed: 3': df.iloc[0]['Unnamed: 3'],'Unnamed: 4': df.iloc[0]['Unnamed: 4'] }).fillna('')
    return net_income


# total_revenue= Revenue_line(df)
# main_expenses = Main_Expenses(df)
# other_expenses = Other_Expenses(df)
# net_income = Net_Income(df)
# main_file = pd.concat([total_revenue, main_expenses, other_expenses, net_income], ignore_index=False, axis=0)


def main():
    
    df = pd.read_excel(r'E:\Python Projects\Grindstone\Variance Report Data - Output Example\Example Input Data\Company 3 Profit & Loss November 2020.xlsx')
    df =df.iloc[1:,]
    total_revenue= Revenue_line(df)
    main_expenses = Main_Expenses(df)
    other_expenses = Other_Expenses(df)
    net_income = Net_Income(df)
    main_file = pd.concat([total_revenue, main_expenses, other_expenses, net_income], ignore_index=False, axis=0)
    main_file.to_excel(r'E:\Python Projects\Grindstone\Variance Report Data - Output Example\Grind_00.xlsx', index=True)
    

if __name__ == "__main__":
    main()
   
        
 

