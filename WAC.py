import pandas as pd

# Load the data from the Excel sheet
file_path = './Data.xlsx'  # Update the path to the Excel file
transactions_df = pd.read_excel(file_path, sheet_name='B2B Transactions')

# Print column names to verify
print("Column names in the Excel file:", transactions_df.columns)

# Print the first few rows of the dataframe to inspect the column names
print("First few rows of the dataframe:")
print(transactions_df.head())

# Assign the actual column names
grade_col = 'Grade'
product_col = 'Product'
air_surface_col = 'Air / Surface'
seller_col = 'Seller'
buyer_col = 'Buyer'
volume_col = 'Volume'
price_unit_col = 'Price / unit'

# Define a function to determine the transaction type based on seller and method of sale
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

# Function to calculate WAC for each product, grade, and transaction type for each company
def calculate_wac_for_company(df, company_id):
    results = {}

    # Determine transaction type for the specific company
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

# Function to calculate WAC for all companies
def calculate_wac_for_all_companies(df):
    all_results = {}
    for company_id in df[seller_col].unique():
        company_results = calculate_wac_for_company(df, company_id)
        all_results[company_id] = company_results
    return all_results

# Run the WAC calculations for all companies
wac_results = calculate_wac_for_all_companies(transactions_df)

# Display the results
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
