import streamlit as st
from varustatistik_formatter import format_varustatistik_external

st.set_page_config(page_title="Varustatistik-till-Forecast", page_icon="📊")

st.title("📊 Varustatistik-formattering")
st.write("Ladda upp en svensk restaurang-Excel så får du en färdig **external_forecast_output.txt**.")

uploaded = st.file_uploader("Välj Excel-fil", type=["xlsx","xls"])

if uploaded:
    txt = format_varustatistik_external(uploaded)     # your own function
    st.success(f"Hittade {txt.count(chr(10)) - 1} rader. Klicka för att ladda ner:")
    st.download_button(
        label="💾 Ladda ner resultat.txt",
        data=txt,
        file_name="external_forecast_output.txt",
        mime="text/plain"
    )
