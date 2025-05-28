import streamlit as st
import pandas as pd
import io

st.title("Estrazione Taglie da Excel")

uploaded_file = st.file_uploader("Carica il file Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    
    # Mostra anteprima con intestazioni leggibili
    st.subheader("Anteprima tabella grezza (riferimenti visivi)")
    df_display = df.copy()
    df_display.index = [f"Riga {i}" for i in df.index]
    df_display.columns = [f"Col {i}" for i in df.columns]
    st.dataframe(df_display, height=500)

    st.subheader("Impostazioni di estrazione")

    row_taglie = st.number_input(
        "Numero di riga con le taglie (⚠️ inserisci 2 se le vedi su 'Riga 1' sopra)",
        min_value=1, value=2
    )
    sku_col = st.number_input("Numero colonna SKU (es. 3 per 'Colore modello')", min_value=0, value=3)
    start_col = st.number_input("Colonna iniziale del range taglie (es. 4 per colonna E)", min_value=0, value=4)
    end_col = st.number_input("Colonna finale del range taglie (es. 30 per colonna AE)", min_value=1, value=30)
    start_row = st.number_input("Riga iniziale del blocco dati (⚠️ inserisci 3 se i dati partono da 'Riga 2')", min_value=1, value=3)
    end_row = st.number_input("Riga finale del blocco dati (es. 19)", min_value=1, value=19)

    include_extra = st.checkbox("Includi una colonna extra (es. prezzo)?")
    col_extra_1 = None
    if include_extra:
        col_extra_1 = st.number_input("Numero colonna extra da includere (es. 58)", min_value=0, value=0)

    if st.button("Estrai dati"):
        try:
            size_labels = df.iloc[row_taglie - 1, start_col:end_col + 1].values
            output_rows = []

            for i in range(start_row - 1, end_row):
                row = df.iloc[i]
                sku = row[sku_col]
                extra_value = row[col_extra_1] if include_extra else None

                for size_label, qty in zip(size_labels, row[start_col:end_col + 1]):
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
            st.success(f"Totale righe estratte: {len(result_df)}")
            st.dataframe(result_df)

            @st.cache_data
            def convert_df(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            excel_bytes = convert_df(result_df)
            st.download_button(
                "Scarica Excel",
                excel_bytes,
                file_name="taglie_estratte.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Errore durante l'elaborazione: {e}")
