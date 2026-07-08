import pandas as pd
from config import JOIN_OPTION

def combine_dataframes(df_list, join=JOIN_OPTION):
    if not df_list:
        raise ValueError("Список DataFrame пуст")
    valid = [df for df in df_list if df is not None]
    if not valid:
        raise ValueError("Нет загруженных DataFrame")
    combined = pd.concat(valid, ignore_index=True, join=join)
    return combined
