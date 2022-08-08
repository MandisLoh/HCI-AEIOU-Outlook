import pandas as pd

df = pd.read_csv('dateonly2.csv')
print(df['Date'].to_list())
df.to_csv('./dateonly2.csv', sep=';', decimal=',')