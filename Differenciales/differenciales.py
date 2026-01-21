import pandas as pd
import numpy as np


df = pd.read_csv('Raw Data\diff.csv')
df = df.iloc[::-1].reset_index(drop=True)
col = df.columns[9]

n = 8

df = df.iloc[: (len(df) // n) * n]

out = pd.DataFrame(
    df[col].to_numpy().reshape(-1, n),
    columns=[" 0°", " 180°", "0°","180°", "0° ", "180° ", " 0° ", " 180° " ]
)

out.to_csv("Data/datos_agrupados.csv", index=False)


