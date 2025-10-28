# Store Sales and Profit Analysis using Python
# --------------------------------------------
# Requirements:
# pip install pandas matplotlib numpy openpyxl seaborn

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

# 1️⃣ Load the Dataset
file_path = r"C:\Users\Sonic\Downloads\SHADOWFOX INTERMEDIATE\Sample - Superstore.csv"  # ✅ Use raw string to avoid unicode errors
df = pd.read_csv(file_path, encoding='ISO-8859-1', parse_dates=['Order Date', 'Ship Date'])

# 2️⃣ Quick Overview
print("Rows:", df.shape[0], " | Columns:", df.shape[1])
print("\nMissing Values:\n", df.isnull().sum())
print("\nColumn Types:\n", df.dtypes)

# 3️⃣ Convert numeric columns and handle missing/invalid data
for col in ['Sales', 'Quantity', 'Discount', 'Profit']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Optional: Drop rows with missing key values
df.dropna(subset=['Sales', 'Profit'], inplace=True)

# 4️⃣ Feature Engineering
df['Order Month'] = df['Order Date'].dt.to_period('M').astype(str)
df['Order Year'] = df['Order Date'].dt.year

# ✅ Safe division: Avoid divide-by-zero in Profit Margin
df['Profit Margin'] = df.apply(
    lambda x: x['Profit'] / x['Sales'] if x['Sales'] != 0 else np.nan, axis=1
)

# 5️⃣ Basic Statistics
print("\nDescriptive Statistics:\n", df[['Sales', 'Quantity', 'Discount', 'Profit']].describe())

# 6️⃣ Sales and Profit by Category/Sub-Category
cat_summary = df.groupby(['Category', 'Sub-Category'])[['Sales', 'Profit']].sum().sort_values('Sales', ascending=False)
print("\nSales & Profit by Category and Sub-Category:\n", cat_summary.head(10))

# 7️⃣ Region-wise Performance
region_summary = df.groupby('Region')[['Sales', 'Profit']].sum().sort_values('Sales', ascending=False)
print("\nSales & Profit by Region:\n", region_summary)

# 8️⃣ Monthly Sales & Profit Trend
monthly = df.set_index('Order Date').resample('M')[['Sales', 'Profit']].sum()

# 9️⃣ Top Products & Customers
top_products = df.groupby('Product Name')[['Sales', 'Profit']].sum().sort_values('Sales', ascending=False).head(10)
top_customers = df.groupby('Customer Name')[['Sales', 'Profit']].sum().sort_values('Sales', ascending=False).head(10)

# 🔟 Negative Profit Orders
neg_profit = df[df['Profit'] < 0]
print(f"\nOrders with negative profit: {len(neg_profit)} ({len(neg_profit)/len(df)*100:.2f}%)")

# 1️⃣1️⃣ Visualization Section

# Line Plot - Monthly Sales
plt.figure(figsize=(10, 4))
plt.plot(monthly.index, monthly['Sales'], marker='o')
plt.title('Monthly Sales Trend')
plt.xlabel('Date')
plt.ylabel('Sales')
plt.grid(True)
plt.tight_layout()
plt.show()

# Line Plot - Monthly Profit
plt.figure(figsize=(10, 4))
plt.plot(monthly.index, monthly['Profit'], marker='o', color='green')
plt.title('Monthly Profit Trend')
plt.xlabel('Date')
plt.ylabel('Profit')
plt.grid(True)
plt.tight_layout()
plt.show()

# Bar Chart - Sales by Category
plt.figure(figsize=(6, 4))
df.groupby('Category')['Sales'].sum().sort_values().plot(kind='barh')
plt.title('Total Sales by Category')
plt.xlabel('Sales')
plt.tight_layout()
plt.show()

# Bar Chart - Profit by Region
plt.figure(figsize=(6, 4))
region_summary['Profit'].plot(kind='bar', color='orange')
plt.title('Profit by Region')
plt.ylabel('Profit')
plt.tight_layout()
plt.show()

# Bar Chart - Top Products by Sales
plt.figure(figsize=(10, 5))
top_products['Sales'].sort_values().plot(kind='barh')
plt.title('Top 10 Products by Sales')
plt.xlabel('Sales')
plt.tight_layout()
plt.show()

# Histogram - Profit Margin Distribution
plt.figure(figsize=(8, 4))
plt.hist(df['Profit Margin'].dropna().clip(-1, 1), bins=30, color='skyblue', edgecolor='black')  # clip extreme values
plt.title('Distribution of Profit Margin')
plt.xlabel('Profit Margin')
plt.ylabel('Count')
plt.tight_layout()
plt.show()

# Optional: Heatmap of Profit by Category & Region
plt.figure(figsize=(8, 5))
pivot = df.pivot_table(values='Profit', index='Category', columns='Region', aggfunc='sum')
sns.heatmap(pivot, annot=True, fmt=".0f", cmap="YlGnBu")
plt.title("Profit Heatmap: Category vs Region")
plt.tight_layout()
plt.show()

# 1️⃣2️⃣ Save Summaries (✅ reset index for multi-index DataFrames)
cat_summary.reset_index().to_csv('summary_category_summary.csv', index=False)
region_summary.reset_index().to_csv('summary_region_summary.csv', index=False)
top_products.reset_index().to_csv('summary_top_products_sales.csv', index=False)
top_customers.reset_index().to_csv('summary_top_customers.csv', index=False)
monthly.to_csv('summary_monthly.csv', index=True)

print("\n✅ Analysis complete. Summary files saved in the current directory.")
