# =============================================================================
# IMPORT DATA
# =============================================================================

import pandas as pd
data = pd.read_excel('sales_data.xlsx')

# =============================================================================
# EXPLORE DATA
# =============================================================================

data.info()
data.describe()
print(data.columns)
print(data.head(5))
print(data.dtypes)

# =============================================================================
# CLEAN DATA
# =============================================================================

data.isnull().sum() # check missing value

# drop row that missing 'Amount' since this is considered an important feature
data_cleaned = data.dropna(subset = ['Amount'])
data_cleaned.isnull().sum()

# =============================================================================
# QUESTION TO ANSWER:
# 1. How many sales have they made with amounts more than 1000
# 2. How many sales have they made that belong to the Category "Tops" and have a Quantity of 3.
# 3. The Total Sales by Category
# 4. Average Amount by Category and Status
# 5. Total Sales by Fulfilment and Shipment Type
# =============================================================================

# Q1:
gt_amount_1000 = data_cleaned[data_cleaned['Amount'] > 1000]
print(gt_amount_1000)

# Q2:
category_Top_and_qty_3 = data_cleaned[(data_cleaned['Category'] == 'Top') & (data_cleaned['Qty'] == 3)]
print(category_Top_and_qty_3)

# Q3:
category_total_amount = data_cleaned.groupby('Category', as_index = False)['Amount'].sum().sort_values('Amount', ascending = False)
print(category_total_amount)

# Q4:
avg_amount_by_category_and_status = data_cleaned.groupby(['Category', 'Status'], as_index = False)['Amount'].mean().sort_values('Amount', ascending = False)
print(avg_amount_by_category_and_status)

# Q5:
total_sale_by_fulfilment_and_shipment = data_cleaned.groupby(['Fulfilment', 'Courier Status'], as_index = False)['Amount'].sum().sort_values('Amount', ascending = False)
total_sale_by_fulfilment_and_shipment.rename(columns={'Courier Status': 'Shipment'}, inplace=True)
print(total_sale_by_fulfilment_and_shipment)

# =============================================================================
# EXPORT DATA
# =============================================================================

# export extracted data to excel files
avg_amount_by_category_and_status.to_excel('average_sales_by_category_and_status.xlsx', index = False)
total_sale_by_fulfilment_and_shipment.to_excel('total_sales_by_fulfilment_and_shipment.xlsx', index = False)
