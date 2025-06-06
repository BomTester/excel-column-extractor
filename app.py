import streamlit as st
import pandas as pd
import io

def col_index_to_excel_letter(index):
    result = ""
    while index >= 0:
        result = chr(index % 26 + ord('A')) + result
        index = index // 26 - 1
    return result

st.title("📄 Excel/CSV Column Extractor with Cell Position")

uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["csv", "xlsx"])
download_format = st.selectbox("Choose download format", ["txt", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, nrows=0)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl', nrows=0)

        columns = df.columns.tolist()
        st.success(f"✅ Number of columns detected: {len(columns)}")

        # Generate letter + column name pairs
        mapped = [(col_index_to_excel_letter(i), col) for i, col in enumerate(columns)]
        st.write("### 🔠 Column Cell Mappings (Excel-style):")
        st.table(pd.DataFrame(mapped, columns=["Excel Cell", "Column Name"]))

        if download_format == "txt":
            content = "\n".join([f"{cell} - {name}" for cell, name in mapped])
            st.download_button("Download TXT", content, file_name="columns.txt")
        else:
            buffer = io.BytesIO()
            pd.DataFrame(mapped, columns=["Excel Cell", "Column Name"]).to_excel(buffer, index=False)
            st.download_button("Download Excel", buffer.getvalue(), file_name="columns.xlsx")
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
