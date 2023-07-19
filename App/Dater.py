import pandas as pd
from querySim import query_result

# query_result to srcexcel.xlsx. 
# Convert dict to a Pandas DataFrame
df = pd.DataFrame(data=query_result)
df.to_excel("srcexcel.xlsx")

# Read vullingen.xlsx and get the MATERIAL IDs
vullingen_df = pd.read_excel('vullingen.xlsx', header=None)
material_ids = vullingen_df.iloc[:, 0].values.tolist() # Get the first column

# Filter srcexcel.xlsx df based on the MATERIAL IDs
df_filtered = df[df['MATERIAL ID'].isin(material_ids)].copy() # Assuming the column name in srcexcel.xlsx is 'MATERIAL ID'

# Ask the user for the suffix and material type
suffix = input("Wat moet er op het einde van titel toegevoegd worden: ")
material_type = input("Typ het MATERIAL TYPE. Kies tussen 'COMMERCIAL', 'JUNCTION', 'PROGRAMME' of 'LIVE RECORD': ")
# In de webapp wordt dit een dropdown menu met vaste keuzes.

# Modify TITLE and MATERIAL TYPE columns
df_filtered.loc[:, 'TITLE'] = df_filtered['TITLE'] + ' ' + suffix
df_filtered.loc[:, 'MATERIAL TYPE'] = material_type

# Convert the filtered DataFrame to list of dicts
filtered_ids = df_filtered.to_dict('records')

# Save the modified DataFrame to destexcel.xlsx
df_filtered.to_excel('destexcel.xlsx')
