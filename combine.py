import pandas as pd

df1 = pd.read_csv("D:\Margo\python\Output v2\protocol_data_2.csv")
df2 = pd.read_csv("D:\Margo\python\Output v1\chrom_data_clean.csv")

df1['ID']=df1['ID'].replace(regex={'КР':'', 'KP':'','/19':''})
df1.ID = df1.ID.astype('int32')
df2.ID = df2.ID.astype('int32')

df_right = pd.merge(df1,df2,on='ID',how='right')
df_right.to_csv('D:\Margo\python\Output v3\chrom_and_protocols_right.csv')

df_left = pd.merge(df1,df2,on='ID',how='left')
df_left.to_csv('D:\Margo\python\Output v3\chrom_and_protocols_left.csv')

df = pd.merge(df1,df2,on='ID',how='inner')
df.to_csv('D:\Margo\python\Output v3\chrom_and_protocols.csv')