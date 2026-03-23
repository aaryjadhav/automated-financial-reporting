import pandas as pd

sales_df = pd.read_csv("data/raw_sales.csv")

print(sales_df.head())

print(sales_df.info())