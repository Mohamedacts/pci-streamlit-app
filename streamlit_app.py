import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("PCI Excel Batch Statistics Extractor")

st.markdown("""
---
**Mohamed Ali**  
Pavement Engineer  
📞 +966581764292
---
""")

# Initialize session state for file uploader key and filename
if 'upload_key' not in st.session_state:
    st.session_state.upload_key = 0
if 'export_name' not in st.session_state:
    st.session_state.export_name = "PCI_Summary"

# Create form for file uploader with clear-on-submit
with st.form("file_upload_form", clear_on_submit=True):
    uploaded_files = st.file_uploader(
        "Upload PCI Excel files", 
        type=["xlsx"], 
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state.upload_key}"
    )
    
    # Add file control buttons
    col1, col2 = st.columns(2)
    with col1:
        submitted = st.form_submit_button("Process Files 🚀")
    with col2:
        if st.form_submit_button("Clear All Files ❌"):
            st.session_state.upload_key += 1
            st.rerun()

# Add filename customization
st.session_state.export_name = st.text_input(
    "Customize export filename:",
    value=st.session_state.export_name,
    help="Do not include .xlsx extension"
)

results = []

if submitted and uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            # Read data with correct skiprows (5 rows) and no headers
            df = pd.read_excel(uploaded_file, sheet_name='PCI', skiprows=5, header=None, engine='openpyxl')
            
            # PCI values in column index 3 (4th column)
            pci_values = pd.to_numeric(df[3], errors='coerce').dropna()
            
            if len(pci_values) == 0:
                st.warning(f"No PCI values found in {uploaded_file.name}")
                continue
                
            results.append({
                'File Name': uploaded_file.name,
                'Average': round(pci_values.mean(), 2),
                'Max': round(pci_values.max(), 2),
                'Min': round(pci_values.min(), 2),
                'Median': round(pci_values.median(), 2)
            })
            
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")

    if results:
        summary = pd.DataFrame(results)
        st.subheader("Summary Table")
        st.dataframe(summary)

        # Export as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 Download Summary",
            data=output.getvalue(),
            file_name=f"{st.session_state.export_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Browser will prompt for save location"
        )
    else:
        st.warning("No valid PCI data found in any uploaded files.")
elif not uploaded_files:
    st.info("Upload PCI Excel files and click 'Process Files' to begin")
