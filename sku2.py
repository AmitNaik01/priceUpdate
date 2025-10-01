import pandas as pd
from woocommerce import API

# WooCommerce API connection
wcapi = API(
    url="https://luxxesims.com",
    consumer_key="ck_ee0139299eb6f20c9293f73095993a1312a6bc88",
    consumer_secret="cs_918d271a1ea54393365dedaf9a1e2111f70ab935",
    version="wc/v3",
    timeout=200
)

# Read Excel
df = pd.read_excel("global.xlsx")

# Clean column names: remove newlines and extra spaces
df.columns = df.columns.str.replace("\n", "").str.strip()

# Confirm columns exist
print("Columns in Excel:", df.columns)

# Loop through each row
for index, row in df.iterrows():
    excel_sku = str(row["API Product code（Auto Activate）"]).strip()
    new_price = str(row["Retail Price USD"]).strip()
    
    # Find product in WooCommerce by SKU
    response = wcapi.get("products", params={"sku": excel_sku}).json()
    
    if response:
        product = response[0]
        wc_sku = product.get("sku")
        old_price = product.get("regular_price")
        product_id = product.get("id")
        
        # Update product price
        update = wcapi.put(f"products/{product_id}", data={"regular_price": new_price}).json()
        updated_price = update.get("regular_price")
        
        print(f"Excel SKU: {excel_sku} | WC SKU: {wc_sku} | Old Price: {old_price} | New Price: {updated_price}")
    else:
        print(f"Excel SKU: {excel_sku} not found in WooCommerce.")
