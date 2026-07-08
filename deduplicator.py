import pandas as pd

def remove_duplicates(df, subset=None):
    initial = len(df)
    df_dedup = df.drop_duplicates(subset=subset, keep='first')
    removed = initial - len(df_dedup)
    return df_dedup, removed
