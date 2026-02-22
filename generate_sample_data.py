"""
Sample Data Generator for Testing Excel Report Automation
Creates realistic business Excel files in input_files/ folder
"""
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import random

# Ensure input folder exists
os.makedirs('input_files', exist_ok=True)

print("🔄 Generating sample Excel files for testing...")

# =============================================================================
# FILE 1: Sales Data (January)
# =============================================================================
print("📊 Creating: Sales_January.xlsx")

sales_jan = pd.DataFrame({
    'Order_ID': [f'ORD-2024-{str(i).zfill(4)}' for i in range(1, 51)],
    'Customer_Name': np.random.choice(['Acme Corp', 'Global Tech', 'SuperMart', 'MegaStore', 'CityShop', 'QuickBuy'], 50),
    'Product': np.random.choice(['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'USB Drive', 'Headphones', 'Webcam'], 50),
    'Quantity': np.random.randint(1, 20, 50),
    'Unit_Price': np.round(np.random.uniform(15, 1200, 50), 2),
    'Order_Date': pd.date_range(start='2024-01-01', periods=50, freq='D'),
    'Sales_Rep': np.random.choice(['Arun', 'Priya', 'Karthik', 'Divya', 'Suresh'], 50),
    'Region': np.random.choice(['North', 'South', 'East', 'West'], 50),
    'Status': np.random.choice(['Completed', 'Pending', 'Shipped'], 50, p=[0.7, 0.2, 0.1])
})

# Calculate Total
sales_jan['Total_Amount'] = sales_jan['Quantity'] * sales_jan['Unit_Price']

# Create multiple sheets
with pd.ExcelWriter('input_files/Sales_January.xlsx', engine='openpyxl') as writer:
    sales_jan.to_excel(writer, sheet_name='Sales_Data', index=False)
    
    # Summary sheet
    summary_jan = pd.DataFrame({
        'Metric': ['Total Orders', 'Total Revenue', 'Average Order Value', 'Top Product'],
        'Value': [
            len(sales_jan),
            f"₹{sales_jan['Total_Amount'].sum():,.2f}",
            f"₹{sales_jan['Total_Amount'].mean():,.2f}",
            sales_jan['Product'].mode()[0]
        ]
    })
    summary_jan.to_excel(writer, sheet_name='Summary', index=False)

print(f"   ✅ Created with {len(sales_jan)} sales records")

# =============================================================================
# FILE 2: Sales Data (February)
# =============================================================================
print("📊 Creating: Sales_February.xlsx")

sales_feb = pd.DataFrame({
    'Order_ID': [f'ORD-2024-{str(i).zfill(4)}' for i in range(51, 101)],
    'Customer_Name': np.random.choice(['Acme Corp', 'Global Tech', 'SuperMart', 'TechWorld', 'Digital Hub', 'CloudMart'], 50),
    'Product': np.random.choice(['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'SSD Drive', 'Headphones', 'Printer'], 50),
    'Quantity': np.random.randint(1, 25, 50),
    'Unit_Price': np.round(np.random.uniform(20, 1500, 50), 2),
    'Order_Date': pd.date_range(start='2024-02-01', periods=50, freq='D'),
    'Sales_Rep': np.random.choice(['Arun', 'Priya', 'Karthik', 'Vijay', 'Anita'], 50),
    'Region': np.random.choice(['North', 'South', 'East', 'West', 'Central'], 50),
    'Status': np.random.choice(['Completed', 'Pending', 'Shipped', 'Cancelled'], 50, p=[0.75, 0.15, 0.08, 0.02])
})

sales_feb['Total_Amount'] = sales_feb['Quantity'] * sales_feb['Unit_Price']

with pd.ExcelWriter('input_files/Sales_February.xlsx', engine='openpyxl') as writer:
    sales_feb.to_excel(writer, sheet_name='Sales_Data', index=False)
    
    summary_feb = pd.DataFrame({
        'Metric': ['Total Orders', 'Total Revenue', 'Average Order Value', 'Top Region'],
        'Value': [
            len(sales_feb),
            f"₹{sales_feb['Total_Amount'].sum():,.2f}",
            f"₹{sales_feb['Total_Amount'].mean():,.2f}",
            sales_feb['Region'].mode()[0]
        ]
    })
    summary_feb.to_excel(writer, sheet_name='Summary', index=False)

print(f"   ✅ Created with {len(sales_feb)} sales records")

# =============================================================================
# FILE 3: HR Employee Data
# =============================================================================
print("📊 Creating: HR_Employees.xlsx")

departments = ['IT', 'Sales', 'Marketing', 'HR', 'Finance', 'Operations']
locations = ['Chennai', 'Bangalore', 'Hyderabad', 'Mumbai', 'Delhi', 'Pune']

hr_data = pd.DataFrame({
    'Employee_ID': [f'EMP-{str(i).zfill(4)}' for i in range(1, 31)],
    'Name': [f'Employee_{i}' for i in range(1, 31)],
    'Department': np.random.choice(departments, 30),
    'Designation': np.random.choice(['Manager', 'Senior', 'Junior', 'Intern', 'Lead'], 30),
    'Salary': np.random.randint(30000, 150000, 30),
    'Join_Date': pd.date_range(start='2020-01-01', periods=30, freq='ME'),
    'Location': np.random.choice(locations, 30),
    'Performance_Rating': np.round(np.random.uniform(2.5, 5.0, 30), 1),
    'Active': np.random.choice([True, False], 30, p=[0.9, 0.1])
})

with pd.ExcelWriter('input_files/HR_Employees.xlsx', engine='openpyxl') as writer:
    hr_data.to_excel(writer, sheet_name='Employee_List', index=False)
    
    # Department wise summary
    dept_summary = hr_data.groupby('Department').agg({
        'Salary': ['mean', 'sum', 'count']
    }).round(2)
    dept_summary.columns = ['Avg_Salary', 'Total_Salary', 'Head_Count']
    dept_summary.reset_index(inplace=True)
    dept_summary.to_excel(writer, sheet_name='Department_Summary', index=False)

print(f"   ✅ Created with {len(hr_data)} employee records")

# =============================================================================
# FILE 4: Inventory Data
# =============================================================================
print("📊 Creating: Inventory_Stock.xlsx")

categories = ['Electronics', 'Office Supplies', 'Furniture', 'Accessories', 'Software']
suppliers = ['TechCorp', 'OfficeMart', 'FurniWorld', 'SupplyHub', 'Digital Solutions']

inventory = pd.DataFrame({
    'SKU': [f'SKU-{str(i).zfill(5)}' for i in range(1, 41)],
    'Product_Name': [f'Product_{i}' for i in range(1, 41)],
    'Category': np.random.choice(categories, 40),
    'Supplier': np.random.choice(suppliers, 40),
    'Stock_Quantity': np.random.randint(0, 500, 40),
    'Unit_Cost': np.round(np.random.uniform(100, 5000, 40), 2),
    'Reorder_Level': np.random.randint(10, 50, 40),
    'Last_Restocked': pd.date_range(start='2023-06-01', periods=40, freq='W'),
    'Warehouse': np.random.choice(['WH-A', 'WH-B', 'WH-C'], 40)
})

# Calculate stock value
inventory['Stock_Value'] = inventory['Stock_Quantity'] * inventory['Unit_Cost']
inventory['Status'] = inventory.apply(
    lambda x: 'Low Stock' if x['Stock_Quantity'] < x['Reorder_Level'] else 'OK', axis=1
)

with pd.ExcelWriter('input_files/Inventory_Stock.xlsx', engine='openpyxl') as writer:
    inventory.to_excel(writer, sheet_name='Current_Stock', index=False)
    
    # Low stock alert
    low_stock = inventory[inventory['Status'] == 'Low Stock']
    if not low_stock.empty:
        low_stock.to_excel(writer, sheet_name='Low_Stock_Alert', index=False)

print(f"   ✅ Created with {len(inventory)} inventory items")

# =============================================================================
# FILE 5: Marketing Campaign Performance
# =============================================================================
print("📊 Creating: Marketing_Campaigns.xlsx")

channels = ['Social Media', 'Email', 'Google Ads', 'LinkedIn', 'Instagram', 'YouTube']
campaigns = ['Summer Sale', 'New Year', 'Diwali Special', 'Product Launch', 'Brand Awareness']

marketing = pd.DataFrame({
    'Campaign_ID': [f'CAMP-{str(i).zfill(3)}' for i in range(1, 21)],
    'Campaign_Name': np.random.choice(campaigns, 20),
    'Channel': np.random.choice(channels, 20),
    'Start_Date': pd.date_range(start='2024-01-01', periods=20, freq='W'),
    'Budget': np.random.randint(10000, 100000, 20),
    'Spend': np.random.randint(5000, 95000, 20),
    'Impressions': np.random.randint(1000, 1000000, 20),
    'Clicks': np.random.randint(50, 50000, 20),
    'Conversions': np.random.randint(5, 5000, 20)
})

# Calculate metrics
marketing['CTR'] = (marketing['Clicks'] / marketing['Impressions'] * 100).round(2)
marketing['CPC'] = (marketing['Spend'] / marketing['Clicks']).round(2)
marketing['ROAS'] = (marketing['Conversions'] * 1000 / marketing['Spend']).round(2)

with pd.ExcelWriter('input_files/Marketing_Campaigns.xlsx', engine='openpyxl') as writer:
    marketing.to_excel(writer, sheet_name='Campaign_Data', index=False)
    
    # Channel performance summary
    channel_perf = marketing.groupby('Channel').agg({
        'Spend': 'sum',
        'Conversions': 'sum',
        'ROAS': 'mean'
    }).round(2)
    channel_perf.reset_index(inplace=True)
    channel_perf.to_excel(writer, sheet_name='Channel_Performance', index=False)

print(f"   ✅ Created with {len(marketing)} campaign records")

# =============================================================================
# Summary
# =============================================================================
print("\n" + "="*60)
print("🎉 SAMPLE FILES CREATED SUCCESSFULLY!")
print("="*60)
print("\n📁 Files in input_files/ folder:")
for file in os.listdir('input_files'):
    if file.endswith('.xlsx'):
        size = os.path.getsize(f'input_files/{file}') / 1024
        print(f"   • {file} ({size:.1f} KB)")
print("\n👉 Now run: python report_generator.py")
print("="*60)
