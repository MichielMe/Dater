import pandas as pd
from querySim import query_result

# query_result naar srcexcel.xlsx 
# Converteer dict naar een Pandas DataFrame
df = pd.DataFrame(data=query_result)
# Bewaar de DataFrame als excel bestand
df.to_excel("srcexcel.xlsx")


# Read vullingen.xlsx and get the MATERIAL IDs
vullingen_df = pd.read_excel('vullingen.xlsx')
material_ids = vullingen_df.iloc[:, 0].values.tolist() # Get the first column

# Filter srcexcel.xlsx based on the MATERIAL IDs
df_filtered = df[df['MATERIAL ID'].isin(material_ids)] # Assuming the column name in srcexcel.xlsx is 'MATERIAL ID'

# Modify TITLE and MATERIAL TYPE columns
df_filtered.loc[:, 'TITLE'] = df_filtered['TITLE'] + ' SEP23'
df_filtered.loc[:, 'MATERIAL TYPE'] = 'COMMERCIAL'


# Convert the filtered DataFrame to list of dicts
filtered_ids = df_filtered.to_dict('records')

# Save the modified DataFrame to destexcel.xlsx
df_filtered.to_excel('destexcel.xlsx')
