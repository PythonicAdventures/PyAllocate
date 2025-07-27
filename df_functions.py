import pandas as pd
import numpy as np

def process_excel_file(file_path):
    """
    Main function that processes an entire Excel file and returns
    a dictionary of processed dataframes for different tabs
    """
    try:
        # Load all sheets from Excel
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        
        # Dictionary to store processed dataframes for each tab
        processed_tabs = {}
        
        # Example: Create different views/tabs from the raw data
        
        # Tab 1: Raw Cap Activity Data
        if 'cap_activity' in excel_data:
            df_contrib = excel_data['cap_activity']
            processed_tabs['Contributions'] = process_contrib(df_contrib)
        
        # Tab 2: Summary by Category
        if 'cap_activity' in excel_data:
            df_reds = excel_data['cap_activity']
            processed_tabs['Redemptions'] = process_reds(df_reds)
        
        # Tab 3: Partner Capital
        if 'partner_capital' in excel_data:
            df_ptr_cap = excel_data['partner_capital']
            processed_tabs['Partner Capital'] = process_partner_alloc(df_ptr_cap)
        return processed_tabs
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return {}
    

def process_contrib(df):
    df_alloc_contrib = (df.pivot_table(index=['fund_name','investor_name', 'classification'],
                             columns='break_period', values='amount', aggfunc='sum', fill_value=0)
                    .query('classification != "redemption"')
                    .reset_index()
                   )
    return df_alloc_contrib

def process_reds(df):   
    df_alloc_red = (df.pivot_table(index=['fund_name','investor_name', 'classification'],
                                columns='break_period', values='amount', aggfunc='sum', fill_value=0)
                        .query('classification != "contribution"')
                        .reset_index()
                    )
    return df_alloc_red


def process_partner_alloc(df):
    
    df_alloc_partner = (df.assign(sub_group_1=df['sub_group_1'].fillna('').astype('category'),
                                sub_group_2=df['sub_group_2'].fillna('').astype('category'),
                                fund_name=df['fund_name'].fillna('').astype('category'),
                                investor_name=df['investor_name'].fillna('').astype('category'),
                                amount=df['amount'].astype(float).fillna(0),
                                break_period=pd.to_datetime(df['break_period']).dt.strftime("%m/%d/%Y")
                                )
                        .pivot_table(index=['fund_name', 'investor_name', 'sub_group_1', 'sub_group_2'],
                                columns='break_period', values='amount', aggfunc='sum', fill_value=0)
                        .reset_index()
                        )
    return df_alloc_partner