import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("PCI Excel Batch Statistics Extractor")

uploaded_files = st.file_uploader(
    "Upload one or more PCI Excel files", 
    type=["xlsx"], 
    accept_multiple_files=True
)

results = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            # Read the PCI sheet, skip first 6 rows (headers)
            df = pd.read_excel(uploaded_file, sheet_name='PCI', skiprows=6, engine='openpyxl')
            pci_col = df.columns[3]
            pci_values = pd.to_numeric(df[pci_col], errors='coerce').dropna()
            if len(pci_values) == 0:
                continue
            avg = pci_values.mean()
            maxv = pci_values.max()
            minv = pci_values.min()
            median = pci_values.median()
            results.append({
                'File Name': uploaded_file.name,
                'Average': avg,
                'Max': maxv,
                'Min': minv,
                'Median': median
            })
        except Exception as e:
            st.warning(f"Error processing {uploaded_file.name}: {e}")

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
        st.info("No valid PCI data found in uploaded files.")
else:
    st.info("Upload PCI Excel files to begin.")

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("PCI Excel Batch Statistics Extractor")

# 👇 Your copyright info
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

# ... rest of your code ...

