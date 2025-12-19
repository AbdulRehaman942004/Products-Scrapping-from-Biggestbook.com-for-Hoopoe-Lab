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

for col in df.columns:
    if 'item' in col.lower() and 'number' in col.lower():
        item_col = col
    if 'link' in col.lower() and 'product' in col.lower():
        link_col = col

if item_col is None or link_col is None:
    print("\nError: Could not find required columns")
    print("Available columns:", df.columns.tolist())
    exit(1)

print(f"\nUsing Item Number column: {item_col}")
print(f"Using Link column: {link_col}")

# Get the first 2 non-null links to understand the pattern
sample_links = []
sample_items = []

for idx, row in df.iterrows():
    if pd.notna(row[link_col]) and str(row[link_col]).strip() != '':
        sample_links.append(str(row[link_col]).strip())
        sample_items.append(str(row[item_col]).strip() if pd.notna(row[item_col]) else '')
        if len(sample_links) >= 2:
            break

print(f"\nSample Item Numbers: {sample_items}")
print(f"Sample Links: {sample_links}")

# Analyze the pattern
if len(sample_links) < 2:
    print("\nError: Need at least 2 sample links to identify the pattern")
    exit(1)

# Try to identify the pattern
# Common patterns: replace item number in URL, append item number, etc.
link1 = sample_links[0]
link2 = sample_links[1]
item1 = sample_items[0]
item2 = sample_items[1]

print(f"\nAnalyzing pattern:")
print(f"Link 1: {link1}")
print(f"Item 1: {item1}")
print(f"Link 2: {link2}")
print(f"Item 2: {item2}")

# Find where item numbers appear in the links
def find_pattern(link1, item1, link2, item2):
    """Try to identify the pattern between item numbers and links"""
    # Check if item number is directly in the link
    if item1 in link1 and item2 in link2:
        # Find the position and pattern
        idx1 = link1.find(item1)
        idx2 = link2.find(item2)
        
        if idx1 == idx2:
            # Same position, likely a direct replacement
            prefix = link1[:idx1]
            suffix = link1[idx1 + len(item1):]
            
            # Verify with second link
            if link2.startswith(prefix) and link2.endswith(suffix):
                return lambda item: prefix + str(item) + suffix
    
    # Check if it's a URL parameter pattern
    # Look for patterns like ?id=ITEM or &item=ITEM or /ITEM/ or /ITEM.html
    patterns = [
        (rf'(\?[^=]*=){re.escape(item1)}([&/]|$)', rf'\1{re.escape(item2)}\2'),
        (rf'(&[^=]*=){re.escape(item1)}([&/]|$)', rf'\1{re.escape(item2)}\2'),
        (rf'/{re.escape(item1)}/', f'/{item2}/'),
        (rf'/{re.escape(item1)}\.html', f'/{item2}.html'),
        (rf'/{re.escape(item1)}$', f'/{item2}'),
    ]
    
    for pattern, replacement in patterns:
        test_link = re.sub(pattern, replacement, link1)
        if test_link == link2:
            return lambda item: re.sub(pattern, lambda m: m.group(1) + str(item) + (m.group(2) if len(m.groups()) > 1 else ''), link1)
    
    # Try simple replacement
    if link1.replace(item1, item2) == link2:
        return lambda item: link1.replace(item1, str(item))
    
    return None

pattern_func = find_pattern(link1, item1, link2, item2)

if pattern_func is None:
    print("\nWarning: Could not automatically detect pattern. Trying manual analysis...")
    # Try a more general approach - find common parts
    common_prefix = ''
    common_suffix = ''
    
    # Find longest common prefix
    for i in range(min(len(link1), len(link2))):
        if link1[i] == link2[i]:
            common_prefix += link1[i]
        else:
            break
    
    # Find longest common suffix
    for i in range(1, min(len(link1), len(link2)) + 1):
        if link1[-i] == link2[-i]:
            common_suffix = link1[-i] + common_suffix
        else:
            break
    
    print(f"Common prefix: {common_prefix}")
    print(f"Common suffix: {common_suffix}")
    
    # Extract the variable part
    if item1 in link1 and item2 in link2:
        # Simple replacement pattern
        pattern_func = lambda item, base_link=link1, base_item=item1: base_link.replace(base_item, str(item))
    else:
        print("\nError: Could not identify pattern. Please check the sample links manually.")
        exit(1)

# Generate links for all rows
print("\nGenerating links for all rows...")
generated_count = 0

for idx, row in df.iterrows():
    # Skip if link already exists
    if pd.notna(row[link_col]) and str(row[link_col]).strip() != '':
        continue
    
    # Get item number
    item_num = row[item_col]
    if pd.isna(item_num):
        continue
    
    item_num = str(item_num).strip()
    
    # Generate link
    try:
        new_link = pattern_func(item_num)
        df.at[idx, link_col] = new_link
        generated_count += 1
    except Exception as e:
        print(f"Error generating link for row {idx}, item {item_num}: {e}")

print(f"Generated {generated_count} new links")

# Save back to the same Excel file
print(f"\nSaving to {file_path}...")
df.to_excel(file_path, index=False)
print("Done! Links have been updated in the Excel file.")

