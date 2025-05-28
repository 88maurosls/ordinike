
import streamlit as st
import pandas as pd

st.title("Estrazione Righe per Taglia da Excel")

uploaded_file = st.file_uploader("Carica il file Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    st.write("Anteprima del file:", df.head(10))

    row_taglie = st.number_input("Numero di riga con le taglie (intestazione)", min_value=1, value=2)
    sku_col = st.number_input("Numero colonna SKU (es. 'Colore modello')", min_value=0, value=3)
    start_col = st.number_input("Colonna iniziale del range taglie (es. colonna E = 4)", min_value=0, value=4)
    end_col = st.number_input("Colonna finale del range taglie (es. colonna AE = 30)", min_value=1, value=30)
    start_row = st.number_input("Riga iniziale del blocco dati", min_value=1, value=3)
    end_row = st.number_input("Riga finale del blocco dati", min_value=1, value=19)

    col_extra_1 = st.number_input("Numero colonna extra da includere (es. prezzo, opzionale)", min_value=0, value=0)
    include_extra = st.checkbox("Includi colonna extra?", value=False)

    if st.button("Estrai dati"):
        size_labels = df.iloc[row_taglie - 1, start_col:end_col+1].values
        output_rows = []

        for i in range(start_row - 1, end_row):
            row = df.iloc[i]
            sku = row[sku_col]
            extra_value = row[col_extra_1] if include_extra else None

            for size_label, qty in zip(size_labels, row[start_col:end_col+1]):
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    data_row = {
                        "SKU": sku,
                        "Size": size_label,
                        "Qty": int(qty)
                    }
                    if include_extra:
                        data_row["Extra"] = extra_value
                    output_rows.append(data_row)

        result_df = pd.DataFrame(output_rows)
        st.write("Risultato:", result_df)

        @st.cache_data
        def convert_df(df):
            return df.to_excel(index=False, engine='openpyxl')

        excel_bytes = convert_df(result_df)
        st.download_button("Scarica Excel", excel_bytes, file_name="taglie_estratte.xlsx")
