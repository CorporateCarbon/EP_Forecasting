from helpers.clean_mi_export import clean_master_inventory_export

import pandas as pd

df = pd.read_excel("Master_Inventory.xlsx")
df_clean = clean_master_inventory_export(df)


