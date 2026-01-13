import pandas as pd
import os
import sys
import re
from datetime import datetime
import shutil
from pathlib import Path

# Define the categories (6 main + 1 anonymous for unmatched products)
CATEGORIES = [
    "Computer Hardware Solutions",
    "Information Technology Services",
    "Office Products & Supplies",
    "Industrial Products & Services",
    "Furniture & Furnishings",
    "Medical Equipment & Supplies",
    "Anonymous"  # For products that don't match any category
]

# Comprehensive keyword mapping for each category
# Keywords are weighted - more specific keywords have higher weights
# Updated based on actual product data analysis
CATEGORY_KEYWORDS = {
    "Computer Hardware Solutions": {
        "high_weight": [
            "computer", "laptop", "desktop", "server", "workstation", "tablet", "pc",
            "cpu", "processor", "motherboard", "ram", "memory", "hard drive", "ssd", "hdd",
            "graphics card", "gpu", "video card", "monitor", "display", "keyboard", "mouse",
            "printer", "scanner", "router", "switch", "network adapter", "usb", "ethernet",
            "wireless", "bluetooth", "webcam", "speaker", "headphone", "microphone",
            "toner", "cartridge", "ink", "printer paper", "laser", "inkjet", "multifunction"
        ],
        "medium_weight": [
            "hardware", "component", "peripheral", "accessory", "cable", "connector",
            "adapter", "dock", "charger", "battery", "power supply", "ups",
            "copier", "fax", "document", "imaging"
        ],
        "low_weight": [
            "electronic", "digital", "tech", "it equipment", "computing", "office equipment"
        ]
    },
    "Information Technology Services": {
        "high_weight": [
            "software", "license", "subscription", "cloud", "saas", "paas", "iaas",
            "support", "maintenance", "consulting", "implementation", "integration",
            "training", "certification", "managed services", "hosting", "backup",
            "security", "firewall", "antivirus", "vpn", "encryption", "authentication"
        ],
        "medium_weight": [
            "service", "solution", "platform", "system", "application", "api",
            "database", "server management", "network management", "it support"
        ],
        "low_weight": [
            "technology", "it", "information technology", "digital transformation"
        ]
    },
    "Office Products & Supplies": {
        "high_weight": [
            "paper", "notebook", "folder", "binder", "envelope", "stamp", "label",
            "pen", "pencil", "marker", "highlighter", "stapler", "staples", "clip",
            "rubber band", "tape", "glue", "scissors", "ruler", "calculator",
            "desk organizer", "file cabinet", "whiteboard", "bulletin board",
            "calendar", "planner", "notepad", "post-it", "sticky note",
            "glue stick", "adhesive", "correction", "whiteout", "eraser",
            "paper clip", "binder clip", "push pin", "thumbtack", "rubber stamp",
            "envelope", "mailing", "shipping label", "file folder", "hanging folder",
            "index card", "note card", "greeting card", "stationery", "letterhead",
            "copy paper", "printer paper", "notebook paper", "graph paper", "legal pad",
            "bathroom", "washroom", "restroom", "soap", "dispenser", "hand soap",
            "paper towel", "toilet paper", "tissue", "napkin", "facial tissue",
            "sanitizer", "hand sanitizer", "purell", "antiseptic", "disinfectant wipes"
        ],
        "medium_weight": [
            "office supply", "stationery", "writing", "filing", "organizer",
            "desk accessory", "bathroom accessory", "washroom accessory",
            "hygiene", "personal care", "cleaning supply", "janitorial",
            "food service", "disposable", "single use", "takeout", "container",
            "plate", "bowl", "cup", "lid", "utensil", "fork", "knife", "spoon",
            "platter", "tray", "coffee", "drink", "hot cup", "cold cup"
        ],
        "low_weight": [
            "office", "supply", "business", "workplace", "commercial", "institutional"
        ]
    },
    "Industrial Products & Services": {
        "high_weight": [
            "industrial", "manufacturing", "machinery", "equipment", "tool",
            "power tool", "hand tool", "drill", "saw", "wrench", "screwdriver",
            "hammer", "pliers", "safety", "ppe", "helmet", "gloves", "goggles",
            "work boot", "safety vest", "hard hat", "welding", "cutting",
            "fastener", "bolt", "nut", "screw", "nail", "rivet", "adhesive",
            "lubricant", "grease", "oil", "paint", "coating", "material",
            "steel", "aluminum", "plastic", "rubber", "fabric",
            "safety glasses", "safety eyewear", "protective eyewear", "goggles",
            "coverall", "coveralls", "protective clothing", "ppe", "personal protective",
            "chemical splash", "particle protection", "liquid protection",
            "safety shoe", "work boot", "steel toe", "safety boot",
            "shoe cover", "shoe covers", "boot cover", "boot covers",
            "sleeve", "sleeves", "arm protection", "protective sleeve",
            "hard hat", "safety helmet", "bump cap", "safety cap",
            "safety vest", "reflective vest", "high visibility", "hi-vis",
            "respirator", "face mask", "dust mask", "n95", "respiratory protection",
            "ear protection", "earplug", "earmuff", "hearing protection",
            "cleaning", "cleaner", "degreaser", "floor cleaner", "concrete cleaner",
            "stripper", "floor stripper", "sealer", "finish", "wax",
            "mop", "mop head", "broom", "broom head", "sweep", "scrub",
            "pad", "floor pad", "scrub pad", "stripping pad", "polishing pad",
            "waste receptacle", "trash can", "garbage can", "receptacle", "container",
            "lock", "keyed lock", "lock core", "security", "access control",
            "hook", "garment hook", "coat hook", "hanger", "rack",
            "tampon dispenser", "sanitary dispenser", "feminine hygiene"
        ],
        "medium_weight": [
            "construction", "maintenance", "repair", "mro", "supply chain",
            "warehouse", "logistics", "packaging", "shipping",
            "janitorial", "custodial", "facility maintenance", "building maintenance",
            "industrial cleaning", "commercial cleaning", "professional cleaning",
            "safety equipment", "protective equipment", "workplace safety"
        ],
        "low_weight": [
            "industrial", "commercial", "professional", "trade", "maintenance"
        ]
    },
    "Furniture & Furnishings": {
        "high_weight": [
            "furniture", "chair", "desk", "table", "cabinet", "shelf", "bookcase",
            "sofa", "couch", "armchair", "ottoman", "coffee table", "end table",
            "dining table", "dining chair", "bed", "mattress", "dresser", "wardrobe",
            "nightstand", "headboard", "futon", "recliner", "office chair",
            "ergonomic", "executive", "task chair", "guest chair", "stool",
            "bench", "counter", "bar stool", "filing cabinet", "storage",
            "pedestal", "credenza", "workstation", "l-shaped", "u-shaped",
            "extension", "desk extension", "work surface", "workstation",
            "seating", "seat", "bath seat", "shower seat", "accessible",
            "rolling", "platform", "step", "rolling step", "step stool",
            "bridge", "desk bridge", "monitor bridge", "keyboard tray"
        ],
        "medium_weight": [
            "furnishing", "decor", "accessory", "lamp", "lighting", "rug", "carpet",
            "curtain", "blind", "shade", "pillow", "cushion", "throw", "blanket",
            "office furniture", "workplace furniture", "commercial furniture"
        ],
        "low_weight": [
            "furniture", "furnishing", "home", "office furniture", "workplace"
        ]
    },
    "Medical Equipment & Supplies": {
        "high_weight": [
            "medical", "hospital", "clinic", "patient", "diagnostic", "monitor",
            "stethoscope", "thermometer", "blood pressure", "sphygmomanometer",
            "defibrillator", "ventilator", "respirator", "oxygen", "suction",
            "surgical", "scalpel", "forceps", "syringe", "needle", "catheter",
            "bandage", "gauze", "dressing", "splint", "brace", "crutch",
            "wheelchair", "walker", "cane", "exam table", "hospital bed",
            "disposable", "sterile", "sanitizer", "disinfectant",
            "medical equipment", "hospital equipment", "clinical equipment",
            "patient care", "medical device", "diagnostic equipment",
            "surgical instrument", "medical instrument", "therapeutic equipment"
        ],
        "medium_weight": [
            "healthcare", "health care", "medical supply", "hospital supply",
            "clinical", "therapeutic", "rehabilitation", "therapy",
            "first aid", "emergency", "ambulance", "medical cart"
        ],
        "low_weight": [
            "medical", "health", "wellness", "care", "hospital", "clinic"
        ]
    }
}

# Weight multipliers
WEIGHTS = {
    "high_weight": 3.0,
    "medium_weight": 2.0,
    "low_weight": 1.0
}

def normalize_text(text):
    """Normalize text for matching"""
    if pd.isna(text) or text == '':
        return ''
    text = str(text).lower()
    # Remove special characters but keep spaces
    text = re.sub(r'[^\w\s]', ' ', text)
    # Normalize whitespace
    text = ' '.join(text.split())
    return text

def calculate_category_score(text, category):
    """Calculate a score for how well text matches a category"""
    if not text:
        return 0.0
    
    text_normalized = normalize_text(text)
    score = 0.0
    
    keywords = CATEGORY_KEYWORDS[category]
    
    # Check each weight level
    for weight_level, weight_multiplier in WEIGHTS.items():
        for keyword in keywords[weight_level]:
            # Count occurrences (case-insensitive)
            count = len(re.findall(r'\b' + re.escape(keyword) + r'\b', text_normalized))
            if count > 0:
                score += count * weight_multiplier
    
    return score

def categorize_product(existing_category, description):
    """Categorize a product based on existing category and description"""
    # Combine existing category and description
    # Give more weight to existing category if it exists
    combined_text = ""
    category_text = ""
    desc_text = ""
    
    if existing_category and str(existing_category).strip():
        category_text = str(existing_category).strip()
        combined_text += category_text + " " + category_text + " "  # Double weight for category
    if description and str(description).strip():
        desc_text = str(description).strip()
        combined_text += desc_text
    
    if not combined_text.strip():
        return "Anonymous", 0.0  # No text to analyze
    
    # Calculate scores for each category (excluding Anonymous)
    category_scores = {}
    main_categories = [cat for cat in CATEGORIES if cat != "Anonymous"]
    
    for category in main_categories:
        score = calculate_category_score(combined_text, category)
        category_scores[category] = score
    
    # Find the category with the highest score
    if not category_scores or max(category_scores.values()) == 0:
        return "Anonymous", 0.0  # No matches found
    
    best_category = max(category_scores, key=category_scores.get)
    best_score = category_scores[best_category]
    
    # Calculate confidence (normalize score)
    total_score = sum(category_scores.values())
    confidence = (best_score / total_score * 100) if total_score > 0 else 0.0
    
    # Boost confidence if best score is significantly higher than second best
    sorted_scores = sorted(category_scores.values(), reverse=True)
    if len(sorted_scores) > 1:
        score_diff = sorted_scores[0] - sorted_scores[1]
        if score_diff > sorted_scores[0] * 0.5:  # If best is 50%+ higher than second
            confidence = min(confidence * 1.2, 100.0)  # Boost confidence by 20% (cap at 100%)
    
    # If confidence is too low, assign to Anonymous
    # Threshold: if confidence < 30% or best_score is very low, mark as Anonymous
    CONFIDENCE_THRESHOLD = 30.0  # Minimum confidence to assign to a category
    MIN_SCORE_THRESHOLD = 2.0     # Minimum absolute score needed
    
    if confidence < CONFIDENCE_THRESHOLD or best_score < MIN_SCORE_THRESHOLD:
        return "Anonymous", confidence
    
    return best_category, confidence

def process_excel_file(excel_path, category_col_name="Category", existing_category_col=None, description_col=None, image_url_col=None):
    """Process the Excel file and categorize products"""
    
    print("="*70)
    print("PRODUCT CATEGORIZATION TOOL")
    print("="*70)
    
    # Read Excel file
    print(f"\nReading Excel file: {excel_path}")
    try:
        if not os.path.exists(excel_path):
            print(f"ERROR: Excel file not found at: {excel_path}")
            return False
        
        # Check if file is accessible
        try:
            test_file = open(excel_path, 'r+b')
            test_file.close()
        except PermissionError:
            print(f"ERROR: Excel file is locked or in use!")
            print("Please close the Excel file if it's open and try again.")
            return False
        
        df = pd.read_excel(excel_path)
        print(f"Successfully loaded Excel file with {len(df)} rows")
        print(f"Columns: {', '.join(df.columns.tolist())}")
        
    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        return False
    
    # Find columns
    print("\n" + "="*70)
    print("Finding columns...")
    
    # Find existing category column
    if existing_category_col:
        if existing_category_col not in df.columns:
            print(f"ERROR: Column '{existing_category_col}' not found")
            return False
        existing_cat_col = existing_category_col
    else:
        # Auto-detect - prioritize "Product Name" as it contains scraped category data
        existing_cat_col = None
        
        # First, check for "Product Name" column (contains scraped category from website)
        for col in df.columns:
            col_lower = col.lower()
            if col_lower == 'product name' or col_lower == 'productname':
                existing_cat_col = col
                break
        
        # Fallback: look for columns with "category" in name (but not the target Category column)
        if not existing_cat_col:
            for col in df.columns:
                col_lower = col.lower()
                if 'category' in col_lower or 'catogary' in col_lower or 'catagory' in col_lower:
                    if col_lower != category_col_name.lower():
                        existing_cat_col = col
                        break
    
    # Find description column
    if description_col:
        if description_col not in df.columns:
            print(f"ERROR: Column '{description_col}' not found")
            return False
        desc_col = description_col
    else:
        # Auto-detect
        desc_col = None
        for col in df.columns:
            col_lower = col.lower()
            if 'description' in col_lower or 'desc' in col_lower:
                desc_col = col
                break
    
    if not desc_col:
        print("ERROR: Could not find description column")
        print("Available columns:", df.columns.tolist())
        return False
    
    # Find Image URL column
    if image_url_col:
        if image_url_col not in df.columns:
            print(f"ERROR: Column '{image_url_col}' not found")
            return False
        img_url_col = image_url_col
    else:
        # Auto-detect
        img_url_col = None
        for col in df.columns:
            col_lower = col.lower()
            if 'image' in col_lower and 'url' in col_lower:
                img_url_col = col
                break
    
    print(f"Existing Category column: {existing_cat_col if existing_cat_col else 'Not found (will use description only)'}")
    print(f"Description column: {desc_col}")
    print(f"Image URL column: {img_url_col if img_url_col else 'Not found'}")
    print(f"Target Category column: {category_col_name}")
    
    # Create or verify target category column
    overwrite = 'n'  # Default: don't overwrite
    if category_col_name not in df.columns:
        df[category_col_name] = ""
        print(f"\nCreated new column: {category_col_name}")
    else:
        # Check how many are already filled
        filled = df[category_col_name].notna() & (df[category_col_name].astype(str).str.strip() != '')
        filled_count = filled.sum()
        print(f"\nColumn '{category_col_name}' already exists with {filled_count} values filled")
        if filled_count > 0:
            overwrite = input(f"Do you want to overwrite existing values? (y/n): ").strip().lower()
            if overwrite != 'y':
                print("Keeping existing values, only filling empty cells...")
    
    # Create backup
    backup_folder = os.path.join(os.path.dirname(excel_path), "Backups")
    os.makedirs(backup_folder, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(backup_folder, f"categorization_backup_{timestamp}.xlsx")
    shutil.copy2(excel_path, backup_path)
    print(f"\nBackup created: {backup_path}")
    
    # Process products
    print("\n" + "="*70)
    print("Processing products...")
    print("="*70)
    
    categorized_count = 0
    skipped_count = 0
    low_confidence_count = 0
    anonymous_count = 0
    not_found_count = 0
    confidence_scores = []
    
    for idx, row in df.iterrows():
        # Skip if already has a value and we're not overwriting
        if category_col_name in df.columns:
            current_value = str(row[category_col_name]).strip() if pd.notna(row[category_col_name]) else ''
            if current_value and current_value != '':
                # Check if we should overwrite
                if overwrite != 'y':
                    skipped_count += 1
                    continue
        
        # Get existing category, description, and image URL
        existing_cat = row[existing_cat_col] if existing_cat_col else None
        desc = row[desc_col] if desc_col else None
        img_url = row[img_url_col] if img_url_col else None
        
        # Check if product was not found (check in Image URL, Product Name, and Description)
        is_not_found = False
        if existing_cat_col and pd.notna(existing_cat):
            if str(existing_cat).strip() == 'Product not found':
                is_not_found = True
        if desc_col and pd.notna(desc):
            if str(desc).strip() == 'Product not found':
                is_not_found = True
        if img_url_col and pd.notna(img_url):
            if str(img_url).strip() == 'Product not found':
                is_not_found = True
        
        # If product was not found, mark Category as "Product not found" and skip categorization
        if is_not_found:
            df.at[idx, category_col_name] = "Product not found"
            not_found_count += 1
            categorized_count += 1
            
            # Show progress every 100 products
            if categorized_count % 100 == 0:
                print(f"Processed {categorized_count} products...")
            continue
        
        # Categorize
        category, confidence = categorize_product(existing_cat, desc)
        
        if category:
            df.at[idx, category_col_name] = category
            confidence_scores.append(confidence)
            
            if category == "Anonymous":
                anonymous_count += 1
            elif confidence < 50:
                low_confidence_count += 1
            
            categorized_count += 1
            
            # Show progress every 100 products
            if categorized_count % 100 == 0:
                avg_confidence = sum(confidence_scores) / len(confidence_scores) if confidence_scores else 0
                print(f"Processed {categorized_count} products... (Avg confidence: {avg_confidence:.1f}%)")
        else:
            skipped_count += 1
        
        # Save progress every 5000 products
        if categorized_count > 0 and categorized_count % 5000 == 0:
            print(f"Saving progress... ({categorized_count} products categorized)")
            df.to_excel(excel_path, index=False)
    
    # Final save
    print(f"\nSaving final results...")
    df.to_excel(excel_path, index=False)
    
    # Statistics
    print("\n" + "="*70)
    print("CATEGORIZATION COMPLETE")
    print("="*70)
    print(f"\nTotal products processed: {len(df)}")
    print(f"Products categorized: {categorized_count}")
    print(f"Products skipped: {skipped_count}")
    print(f"Products marked as 'Product not found': {not_found_count}")
    print(f"Products marked as Anonymous: {anonymous_count}")
    print(f"Low confidence (<50%): {low_confidence_count}")
    
    if confidence_scores:
        avg_confidence = sum(confidence_scores) / len(confidence_scores)
        print(f"Average confidence: {avg_confidence:.1f}%")
        print(f"Min confidence: {min(confidence_scores):.1f}%")
        print(f"Max confidence: {max(confidence_scores):.1f}%")
    
    # Category distribution
    print("\nCategory Distribution:")
    if category_col_name in df.columns:
        category_counts = df[category_col_name].value_counts()
        for cat, count in category_counts.items():
            if pd.notna(cat) and str(cat).strip():
                print(f"  {cat}: {count}")
    
    # Show low confidence products for review
    if low_confidence_count > 0:
        print(f"\n⚠️  {low_confidence_count} products have low confidence (<50%)")
        print("Consider reviewing these manually for accuracy.")
    
    # Show product not found info
    if not_found_count > 0:
        print(f"\n❌ {not_found_count} products were marked as 'Product not found'")
        print("These products were not found on the website during scraping.")
    
    # Show Anonymous products info
    if anonymous_count > 0:
        print(f"\n📋 {anonymous_count} products were marked as 'Anonymous'")
        print("These products didn't match any category well enough.")
        print("You may want to review these and manually assign categories if needed.")
    
    print(f"\nResults saved to: {excel_path}")
    print(f"Backup saved to: {backup_path}")
    
    return True

def main():
    """Main function"""
    # Get Excel file path
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        # Default: look for Excel file in parent directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(script_dir)
        excel_path = os.path.join(parent_dir, "ScrappedProducts.xlsx")
        
        if not os.path.exists(excel_path):
            excel_path = input("Enter path to Excel file: ").strip().strip('"')
    
    # Column names (can be customized)
    category_col = input("Enter target Category column name (default: 'Category'): ").strip()
    if not category_col:
        category_col = "Category"
    
    existing_cat_col = input("Enter existing category column name (press Enter to auto-detect): ").strip()
    if not existing_cat_col:
        existing_cat_col = None
    
    desc_col = input("Enter description column name (press Enter to auto-detect): ").strip()
    if not desc_col:
        desc_col = None
    
    # Process the file
    success = process_excel_file(excel_path, category_col, existing_cat_col, desc_col)
    
    if success:
        print("\n✅ Categorization completed successfully!")
    else:
        print("\n❌ Categorization failed. Please check the errors above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
