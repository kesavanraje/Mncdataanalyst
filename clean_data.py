
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent
DATA = BASE / "data"

raw_file = DATA / "raw_sales_data.csv"
clean_file = DATA / "cleaned_sales_data.csv"

# Load raw dataset
df = pd.read_csv(raw_file)

print("BEFORE CLEANING")
print("Rows:", len(df))
print("Duplicate rows:", df.duplicated().sum())
print("Missing values:\n", df.isnull().sum())

# Remove duplicate records
df = df.drop_duplicates()

# Standardize text columns
for col in ["City", "State", "Region", "Segment",
            "Product", "Category", "Subcategory",
            "CustomerName", "PaymentMethod"]:
    df[col] = df[col].astype("string").str.strip()

# Standardize segment names
df["Segment"] = df["Segment"].str.title()

# Fix known state-name variants
df["State"] = df["State"].replace({
    "Maharashtra": "Maharashtra"
})

# Fill missing categories using product information
category_map = {
    "Laptop": ("Electronics", "Computers"),
    "Mouse": ("Electronics", "Accessories"),
    "Keyboard": ("Electronics", "Accessories"),
    "Monitor": ("Electronics", "Displays"),
    "Office Chair": ("Furniture", "Chairs"),
    "Desk": ("Furniture", "Tables"),
    "Notebook": ("Office Supplies", "Paper"),
    "Pen Pack": ("Office Supplies", "Writing"),
    "Printer": ("Electronics", "Printers"),
    "Bookshelf": ("Furniture", "Storage")
}

for index, row in df.iterrows():
    product = row["Product"]
    if product in category_map:
        category, subcategory = category_map[product]
        if pd.isna(row["Category"]) or not str(row["Category"]).strip():
            df.at[index, "Category"] = category
        if pd.isna(row["Subcategory"]) or not str(row["Subcategory"]).strip():
            df.at[index, "Subcategory"] = subcategory

# Fill missing cities with a transparent placeholder
df["City"] = df["City"].fillna("Unknown")

# Validate key fields
df = df.dropna(subset=["OrderID", "OrderDate", "CustomerID"])
df["OrderDate"] = pd.to_datetime(df["OrderDate"], errors="coerce")
df = df.dropna(subset=["OrderDate"])

# Save cleaned data
df.to_csv(clean_file, index=False, encoding="utf-8-sig")

print("\nAFTER CLEANING")
print("Rows:", len(df))
print("Duplicate rows:", df.duplicated().sum())
print("Saved to:", clean_file)
