# main.py

from market_cap_utils import extract_market_cap, create_market_cap_slide
import os

# Define the 3Q data directory
DATA_DIR = os.path.join("data", "Q1")

# Gather all model files that end with "Model 9-30-24.xlsx"
excel_files = [
    os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR)
    if f.endswith("Model 9-30-24.xlsx")
]

# Find the 3Q PowerPoint template file
pptx_input = os.path.join(DATA_DIR, "Cat Rock Capital 3Q24 Review Webinar Presentation - Technical Case.pptx")
pptx_output = os.path.join(DATA_DIR, "Updated_MarketCap_Slide.pptx")

# Extract market cap for each Excel model
market_cap_data = [extract_market_cap(file) for file in excel_files]

# Create PowerPoint slide
create_market_cap_slide(
    data=market_cap_data,
    pptx_path=pptx_input,
    output_path=pptx_output
)

print(f"✅ Updated PowerPoint saved to: {pptx_output}")
