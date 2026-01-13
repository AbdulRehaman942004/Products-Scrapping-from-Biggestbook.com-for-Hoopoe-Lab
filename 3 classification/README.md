# Product Categorization Tool

Automatically categorizes products into 6 predefined categories using intelligent keyword matching.

## Categories

1. Computer Hardware Solutions
2. Information Technology Services
3. Office Products & Supplies
4. Industrial Products & Services
5. Furniture & Furnishings
6. Medical Equipment & Supplies

## Features

- **Intelligent Keyword Matching**: Uses weighted keyword matching with high/medium/low priority keywords
- **Confidence Scoring**: Provides confidence scores for each categorization
- **Auto-detection**: Automatically finds category and description columns
- **Progress Tracking**: Shows progress and saves every 500 products
- **Backup System**: Creates automatic backups before processing
- **Low Confidence Flagging**: Identifies products with low confidence for manual review

## Requirements

```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage (Auto-detect columns)

```bash
python categorize_products.py
```

The script will:
1. Look for `ScrappedProducts.xlsx` in the parent directory
2. Auto-detect category and description columns
3. Create/update a "Category" column with categorizations

### Custom Usage

```bash
python categorize_products.py "path/to/your/file.xlsx"
```

Then follow the prompts to:
- Specify target category column name
- Specify existing category column (or press Enter to auto-detect)
- Specify description column (or press Enter to auto-detect)

## How It Works

1. **Reads Excel File**: Loads the Excel file and detects columns
2. **Keyword Matching**: 
   - Combines existing category and description text
   - Matches against weighted keyword lists for each category
   - High-weight keywords: 3x multiplier
   - Medium-weight keywords: 2x multiplier
   - Low-weight keywords: 1x multiplier
3. **Scoring**: Calculates a score for each category
4. **Selection**: Chooses the category with the highest score
5. **Confidence**: Calculates confidence as percentage of total score
6. **Saves Results**: Updates Excel file with categorizations

## Accuracy

- **High Confidence (>70%)**: Very likely correct
- **Medium Confidence (50-70%)**: Probably correct, but review recommended
- **Low Confidence (<50%)**: Review manually

## Customization

You can customize the keyword lists in `categorize_products.py`:

```python
CATEGORY_KEYWORDS = {
    "Your Category": {
        "high_weight": ["keyword1", "keyword2"],
        "medium_weight": ["keyword3"],
        "low_weight": ["keyword4"]
    }
}
```

## Output

The script provides:
- Total products processed
- Products categorized
- Average confidence score
- Category distribution
- List of low-confidence products for review

## Notes

- The script creates backups automatically
- Progress is saved every 500 products
- Existing values are preserved unless you choose to overwrite
- Low confidence products are flagged for manual review
