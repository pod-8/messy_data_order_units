# Import packages required
import pandas as pd

# Import messy data from Excel and assign to the mesy_data variabe as a data frame
messy_data = pd.read_excel('messy_order_data.xlsx', sheet_name='Sheet1')

# Cleaning step 1: Remove ',' from the values and replace with ''
messy_data['Price Per Unit (£s)'] = messy_data['Price Per Unit (£s)'].str.replace(',','')

# Cleaning Step 2: Remove '£---' from the values and replace with ''
messy_data['Price Per Unit (£s)'] = messy_data['Price Per Unit (£s)'].str.replace('£---','')

# Cleaning Step 3: Coerce Price Per Unit to be a float 
messy_data['Price Per Unit (£s)'] = messy_data['Price Per Unit (£s)'].astype(float)

# Export data back to excel (index = False ensures the pandas indicies are not exported to excel)
messy_data.to_excel('clean_order_data.xlsx',index=False)