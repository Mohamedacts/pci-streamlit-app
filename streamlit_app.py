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

uploaded_files = st.file_uploader(
    "Upload one or more PCI Excel files", 
    type=["xlsx"], 
    accept_multiple_files=True
)

results = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            # Read data with correct skiprows (5 rows) and no headers
            df = pd.read_excel(uploaded_file, sheet_name='PCI', skiprows=5, header=None, engine='openpyxl')
            
            # PCI values are in column index 3 (4th column)
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
            label="Download Summary as Excel",
            data=output.getvalue(),
            file_name="PCI_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No valid PCI data found in any uploaded files.")
else:
    st.info("Upload PCI Excel files to begin.")
