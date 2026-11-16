import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Usklađivanje Overseas faktura i Izlaza", layout="wide")

st.title("📦 Usklađivanje Overseas faktura i Analize izlaza")
st.markdown("Ova aplikacija spaja podatke s Overseas fakture i vaše interne analize izlaza po broju otpremnice te prikazuje neusklađene stavke.")

# Sidebar za učitavanje datoteka
st.sidebar.header("1. Učitavanje datoteka")
file_overseas = st.sidebar.file_uploader("Učitaj Overseas prilog (Excel)", type=["xlsx", "xls"])
file_analiza = st.sidebar.file_uploader("Učitaj Analizu izlaza (Excel)", type=["xlsx", "xls"])

def extract_ot_number(ref):
    if pd.isna(ref):
        return None
    match = re.search(r'OT\s*(\d+)', str(ref), re.IGNORECASE)
    if match:
        return int(match.group(1))
    clean_str = re.sub(r'\D', '', str(ref))
    if clean_str.isdigit() and len(clean_str) >= 4:
        return int(clean_str)
    return None

if file_overseas and file_analiza:
    @st.cache_data
    def load_data(f_overseas, f_analiza):
        df_ov = pd.read_excel(f_overseas)
        df_an = pd.read_excel(f_analiza, header=1)
        return df_ov, df_an

    df_ov, df_an = load_data(file_overseas, file_analiza)

    # Ekstrakcija broja otpremnice iz Overseas reference (Reference 1)
    df_ov['Extracted_OT'] = df_ov['Reference 1'].apply(extract_ot_number)
    
    # Priprema broja otpremnice u analizi izlaza
    if 'Broj' in df_an.columns:
        df_an['Broj_int'] = pd.to_numeric(df_an['Broj'], errors='coerce')
    else:
        st.error("U Analizi izlaza nije pronađen stupac 'Broj'. Provjerite strukturu tablice.")
        st.stop()

    # Spajanje tablica (Left join s Overseas na Analizu)
    merged_df = pd.merge(
        df_ov,
        df_an,
        left_on='Extracted_OT',
        right_on='Broj_int',
        how='left',
        indicator=True
    )

    # Definiranje uspješno spojenih i neuspjelih
    matched = merged_df[merged_df['_merge'] == 'both']
    unmatched = merged_df[merged_df['_merge'] == 'left_only']

    # Prikaz metrika
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ukupno stavki na fakturi", len(df_ov))
    col2.metric("Uspješno povezano", len(matched), delta=f"{len(matched)/len(df_ov)*100:.1f}%")
    col3.metric("Nije uspjelo povezati", len(unmatched), delta=f"-{len(unmatched)/len(df_ov)*100:.1f}%", delta_color="inverse")
    
    total_cost_unmatched = unmatched['Total Price'].sum() if 'Total Price' in unmatched.columns else 0
    col4.metric("Iznos nepovezanih troškova", f"{total_cost_unmatched:,.2f} EUR")

    # Tabs za pregled
    tab1, tab2, tab3 = st.tabs(["❌ Nepovezane stavke (Faktura)", "✅ Uspješno povezane", "📊 Sveukupni pregled spajanja"])

    with tab1:
        st.subheader("Stavke s Overseas fakture koje se nisu mogle povezati s otpremnicama")
        st.markdown("Ovdje se nalaze pošiljke za koje u polju `Reference 1` nije pronađen odgovarajući broj otpremnice ili otpremnica ne postoji u Analizi izlaza za taj mjesec.")
        
        if len(unmatched) > 0:
            display_cols_ov = [c for c in ['Invoice Number', 'Shipping Date', 'Consignee Name', 'Consignee Town', 'Reference 1', 'Transport Price', 'Total Price'] if c in unmatched.columns]
            st.dataframe(unmatched[display_cols_ov], use_container_width=True)

            csv = unmatched.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Preuzmi nepovezane stavke kao CSV",
                data=csv,
                file_name="nepovezane_overseas_stavke.csv",
                mime="text/csv",
            )
        else:
            st.success("Sve stavke s fakture su uspješno povezane!")

    with tab2:
        st.subheader("Uspješno povezane stavke")
        if len(matched) > 0:
            display_cols_matched = [c for c in ['Invoice Number', 'Extracted_OT', 'Consignee Name', 'Naziv korisnika', 'Transport Price', 'Total Price'] if c in matched.columns]
            st.dataframe(matched[display_cols_matched], use_container_width=True)
        else:
            st.info("Nema uspješno povezanih stavki.")

    with tab3:
        st.subheader("Detaljna tablica spajanja (Merge Result)")
        # Ispravljeno sa 'merged' na 'merged_df'
        st.dataframe(merged_df[['Invoice Number', 'Reference 1', 'Extracted_OT', 'Consignee Name', '_merge']], use_container_width=True)

else:
    st.info("👈 Molimo vas da učitate obje Excel datoteke u bočnoj traci (sidebaru) kako bi aplikacija pokrenula analizu.")
