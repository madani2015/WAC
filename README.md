
# Weighted Average Cost (WAC) Calculation

This script calculates the Weighted Average Cost (WAC) for different grades and transaction types (Air In, Air Out, Surface In, Surface Out) for each company based on transaction data from an Excel file.


Place your data file (`Data.xlsx`) in the same directory as the script.

## Running the Script

To run the script, use the following command:

```sh
python WAC.py
```

The script reads transaction data from `Data.xlsx`, processes it, and calculates the Weighted Average Cost (WAC) for each grade and transaction type.

## Description

The script performs the following steps:

1. Reads data from `Data.xlsx`.
2. Assigns the appropriate column names to variables.
3. Calculates the Weighted Average Cost (WAC) for each grade and each transaction type (Air In, Air Out, Surface In, Surface Out).
4. For each grade and each transaction type, calculates the total volume and total value.
5. Computes the WAC as total value divided by total volume.
6. Groups the transactions by grade, air/surface, seller, and buyer.
7. Calculates total volume and value for each group.
8. Determines transaction types based on whether the company is a seller or buyer:
   - If the company is a seller and sells a product by air, it is categorized as Air Out.
   - If the company is a seller and sells a product by surface, it is categorized as Surface Out.
   - If the company is a buyer and buys a product by air, it is categorized as Air In.
   - If the company is a buyer and buys a product by surface, it is categorized as Surface In.
9. Uses the segregated transaction data to compute the WAC for each grade and transaction type.

## Script Overview

### Loading Data

The script starts by loading data from the specified Excel file:

```python
transactions_df = pd.read_excel(file_path, sheet_name='B2B Transactions')
```

### Assign Column Names

The following variables are used to reference the columns:

```python
grade_col = 'Grade'
product_col = 'Product'
air_surface_col = 'Air / Surface'
seller_col = 'Seller'
buyer_col = 'Buyer'
volume_col = 'Volume'
price_unit_col = 'Price / unit'
```

### Determine Transaction Type

A function to determine the transaction type based on the seller and method of sale:

```python
def determine_transaction_type(row, company_id=1):
    if row[seller_col] == company_id:
        if row[air_surface_col] == 'Air':
            return 'Air Out'
        elif row[air_surface_col] == 'Surface':
            return 'Surface Out'
    elif row[buyer_col] == company_id:
        if row[air_surface_col] == 'Air':
            return 'Air In'
        elif row[air_surface_col] == 'Surface':
            return 'Surface In'
    return 'Unknown'
```

### Calculate WAC for Each Company

A function to calculate the WAC for each product, grade, and transaction type for each company:

```python
def calculate_wac_for_company(df, company_id):
    results = {}
    df['Transaction Type'] = df.apply(lambda row: determine_transaction_type(row, company_id), axis=1)
    for grade in df[grade_col].unique():
        grade_df = df[df[grade_col] == grade]
        for product in grade_df[product_col].unique():
            product_df = grade_df[grade_df[product_col] == product]
            for transaction_type in ['Air In', 'Surface In', 'Air Out', 'Surface Out']:
                transaction_df = product_df[product_df['Transaction Type'] == transaction_type]
                total_volume = transaction_df[volume_col].sum()
                total_value = (transaction_df[volume_col] * transaction_df[price_unit_col]).sum()
                wac = total_value / total_volume if total_volume > 0 else 0
                if grade not in results:
                    results[grade] = {}
                if product not in results[grade]:
                    results[grade][product] = {}
                results[grade][product][transaction_type] = {
                    'Total Volume': total_volume,
                    'Total Value': total_value,
                    'WAC': wac
                }
    return results
```

### Calculate WAC for All Companies

A function to calculate the WAC for all companies:

```python
def calculate_wac_for_all_companies(df):
    all_results = {}
    for company_id in df[seller_col].unique():
        company_results = calculate_wac_for_company(df, company_id)
        all_results[company_id] = company_results
    return all_results
```

### Display Results

The script prints the WAC results for each company:

```python
wac_results = calculate_wac_for_all_companies(transactions_df)
print("\nWeighted Average Costs (WAC) and Inventory Updates for each company:")
for company_id, company_data in wac_results.items():
    print(f"\nCompany ID: {company_id}")
    for grade, product_data in company_data.items():
        for product, transaction_data in product_data.items():
            for transaction_type, metrics in transaction_data.items():
                print(f"Grade: {grade}, Product: {product}, {transaction_type}")
                print(f"  Total Volume: {metrics['Total Volume']}")
                print(f"  Total Value: {metrics['Total Value']:.2f}")
                print(f"  WAC: {metrics['WAC']:.2f}")
```

### Future Enhancements
Implement functions for calculating the Cost of Goods Sold (COGS)
calculating the Cost of Goods Sold (COGS) using the resulting weighted average cost (WAC). The steps for COGS calculation are:

Ending Inventory Q-1
Air Out
Air In
Downgrading Effect1
Surface Out
Use for Production
B2C Sales
Production X
Surface In
Downgrading Effect2

the resulting WAC will be used in Air out, Air In, Surface Out and Surface In in the calculation of COGS. 
For B2C Sales, the entry should be extracted from Sales tab of Data.xls.
For Production cost, a function should be defined as cost of prouction depends on several factors.
Downgrading function needs to be implemented as well.
