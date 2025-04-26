import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Create output folder if it doesn't exist
os.makedirs("output_files", exist_ok=True)

# Load dataset (You can replace the path with your local path)
data = pd.read_csv('Superstore_Sales.csv')

# Data Cleaning
data.dropna(inplace=True)
data['Order Date'] = pd.to_datetime(data['Order Date'])

# Exploratory Data Analysis
# Top 5 cities by total sales
top_cities = data.groupby('City')['Sales'].sum().sort_values(ascending=False).head(5)

# Top product categories by sales
top_categories = data.groupby('Category')['Sales'].sum().sort_values(ascending=False)

# Monthly sales trend
data['Month'] = data['Order Date'].dt.to_period('M')
monthly_sales = data.groupby('Month')['Sales'].sum()

# Profit vs Sales Scatter Plot
plt.figure(figsize=(8,6))
sns.scatterplot(x='Sales', y='Profit', data=data)
plt.title('Profit vs Sales')
plt.xlabel('Sales')
plt.ylabel('Profit')
plt.savefig('output_files/profit_vs_sales.png')  # Save in output_files folder
plt.close()

# Category-wise Profit Margins
category_profit = data.groupby('Category')['Profit'].sum().sort_values(ascending=False)

# Save key outputs to Excel
with pd.ExcelWriter('output_files/Sales_Analysis_Report.xlsx') as writer:  # Save in output_files folder
    top_cities.to_excel(writer, sheet_name='Top Cities')
    top_categories.to_excel(writer, sheet_name='Top Categories')
    monthly_sales.to_frame('Monthly Sales').to_excel(writer, sheet_name='Monthly Sales')
    category_profit.to_excel(writer, sheet_name='Category Profit')

print("Analysis completed! Excel and chart file generated.")
