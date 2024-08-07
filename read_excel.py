#%%
import pandas as pd

watch_df = pd.read_excel('類股清單.xlsx')
column_names = [col_name for col_name in watch_df.columns if 'Unnamed' not in col_name]
table_dict = {}

for i, col_name in enumerate(column_names):
    table_dict[col_name] = watch_df.iloc[:, (2*i):(2*i+2)]
    table_dict[col_name].columns = ['代碼', '名稱']
    table_dict[col_name] = table_dict[col_name].iloc[1:, :]
    table_dict[col_name] = table_dict[col_name].dropna(axis=0, how = 'all')
    table_dict[col_name] = table_dict[col_name].reset_index(drop=True)
# %%
