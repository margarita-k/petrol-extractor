import pandas as pd

df1 = pd.read_csv("protocol_data_2019.csv")
df2 = pd.read_csv("chrom_data_clean.csv")
df = pd.merge(df1,df2,on='ID')
