# app.py

import streamlit as st
import os
import datetime
from market_cap_utils import extract_market_cap, create_market_cap_slide

st.set_page_config(page_title="Market Cap to PowerPoint", layout="centered")
st.title("📊 Market Capitalization Slide Generator")

# Define data location
DATA_DIR = os.path.join("data", "3Q")
pptx_template = os.path.join(DATA_DIR, "Cat Rock Capital 3Q24 Review Webinar Presentation - Technical Case.pptx")

st.markdown("### Step 1: Auto-load 3Q Excel Models")

excel_files = [
    os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR)
    if f.endswith("Model 9-30-24.xlsx")
]

if not excel_files:
    st.warning("No Excel model files found in `data/3Q/`")
else:
    st.success(f"Found {len(excel_files)} model files.")

    with st.expander("Show loaded files"):
        for file in excel_files:
            st.text(os.path.basename(file))

    if st.button("▶ Generate PowerPoint Slide"):
        market_cap_data = [extract_market_cap(f) for f in excel_files]

        # Create output filename once
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(DATA_DIR, f"Updated_MarketCap_Slide_{timestamp}.pptx")

        create_market_cap_slide(market_cap_data, pptx_template, output_path)

        st.success("✅ Slide created successfully!")

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Download Updated PowerPoint",
                data=f,
                file_name=os.path.basename(output_path)
            )

# import streamlit as st
# import os
# import datetime
# from market_cap_utils import update_mkt_cap_column

# st.set_page_config(page_title="Market Cap to PowerPoint", layout="centered")
# st.title("📊 Market Capitalization Slide Generator")

# # Set up file paths
# MODEL_DIR = "data/3Q"

# PPTX_TEMPLATE = "data/3Q/Cat Rock Capital 3Q24 Review Webinar Presentation - Technical Case.pptx"

# timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
# OUTPUT_PPTX = f"data/Updated_MarketCap_{timestamp}.pptx"

# # Find model files
# model_files = [f for f in os.listdir(MODEL_DIR) if f.endswith(".xlsx")]

# if not model_files:
#     st.warning("⚠️ No Excel model files found in /data/models/")
# else:
#     st.success(f"✅ Found {len(model_files)} model file(s)")
#     with st.expander("Show model files"):
#         for f in model_files:
#             st.text(f)

#     if st.button("▶ Generate PowerPoint"):
#         update_mkt_cap_column(PPTX_TEMPLATE, MODEL_DIR, OUTPUT_PPTX)
#         st.success("✅ PowerPoint updated!")

#         with open(OUTPUT_PPTX, "rb") as f:
#             st.download_button(
#                 label="📥 Download Updated PPTX",
#                 data=f,
#                 file_name=os.path.basename(OUTPUT_PPTX)
#             )
