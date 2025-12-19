import pandas as pd
import re
import os

# Read the Excel file (Excel file is in parent folder, script is in subfolder)
file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "ScrappedProducts.xlsx")
df = pd.read_excel(file_path)

# Display column names to verify
print("Column names:", df.columns.tolist())
print("\nFirst few rows:")
print(df.head())

# Find the columns (handle potential variations in naming)
item_col = None
link_col = None

# First, try to find exact "Item Number" column (prioritize this)
for col in df.columns:
    if col.lower() == 'item number':
        item_col = col
        break

# If not found, look for columns containing "item" and "number" but exclude "stock" and "butted"
if item_col is None:
    for col in df.columns:
        col_lower = col.lower()
        if 'item' in col_lower and 'number' in col_lower:
            # Exclude "Item Stock Number-Butted" and similar columns
            if 'stock' not in col_lower and 'butted' not in col_lower:
                item_col = col
                break

# If still not found, use any column with "item" and "number" (fallback)
if item_col is None:
    for col in df.columns:
        if 'item' in col.lower() and 'number' in col.lower():
            item_col = col
            break

# Find link column
for col in df.columns:
    if 'link' in col.lower() and 'product' in col.lower():
        link_col = col
        break

if item_col is None or link_col is None:
    print("\nError: Could not find required columns")
    print("Available columns:", df.columns.tolist())
    exit(1)

print(f"\nUsing Item Number column: {item_col}")
print(f"Using Link column: {link_col}")

# Get the base URL pattern from existing links
base_url = None
for idx, row in df.iterrows():
    if pd.notna(row[link_col]) and str(row[link_col]).strip() != '':
        link = str(row[link_col]).strip()
        # Extract base URL pattern: everything before the itemId parameter value
        # Pattern: https://www.biggestbook.com/ui#/itemDetail?itemId=
        match = re.search(r'(.+?itemId=)', link)
        if match:
            base_url = match.group(1)
            print(f"Found base URL pattern: {base_url}")
            break

# If no existing link found, use the standard pattern
if base_url is None:
    base_url = "https://www.biggestbook.com/ui#/itemDetail?itemId="
    print(f"Using default base URL pattern: {base_url}")

# Function to generate link with exact Item Number
def generate_link(item_number):
    """Generate link using exact Item Number"""
    return base_url + str(item_number).strip()

# Generate links for all rows using exact Item Number
print("\nGenerating links for all rows using exact Item Numbers...")
generated_count = 0
updated_count = 0

for idx, row in df.iterrows():
    # Get item number
    item_num = row[item_col]
    if pd.isna(item_num):
        continue
    
    item_num = str(item_num).strip()
    
    # Generate correct link with exact Item Number
    correct_link = generate_link(item_num)
    
    # Check if link needs to be updated
    current_link = row[link_col]
    if pd.isna(current_link) or str(current_link).strip() == '':
        # No link exists, add it
        df.at[idx, link_col] = correct_link
        generated_count += 1
    else:
        # Link exists, check if it's correct
        current_link = str(current_link).strip()
        if current_link != correct_link:
            # Link is incorrect, update it
            df.at[idx, link_col] = correct_link
            updated_count += 1
            print(f"  Updated row {idx + 1}: {item_num} - {current_link} -> {correct_link}")

print(f"\nGenerated {generated_count} new links")
print(f"Updated {updated_count} incorrect links")
print(f"Total links processed: {generated_count + updated_count}")

# Save back to the same Excel file
print(f"\nSaving to {file_path}...")
df.to_excel(file_path, index=False)
print("Done! Links have been updated in the Excel file.")

