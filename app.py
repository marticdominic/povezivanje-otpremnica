import streamlit as st
import pandas as pd
import re
import difflib

st.set_page_config(page_title="Usklađivanje Overseas faktura i Izlaza", layout="wide")

st.title("📦 Pametno usklađivanje Overseas faktura i Izlaza")
st.markdown("Aplikacija spaja po broju otpremnice, a preostale stavke povezuje po preciznom nazivu kupca iz tvoje tablice!")

# Sidebar za učitavanje datoteka
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

def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).lower()
    for word in ['d.o.o.', 'd.d.', 'j.d.o.o.', '-', 'isporuka', 'skladište', 'doo', 'dd']:
        name = name.replace(word, '')
    return ''.join(e for e in name if e.isalnum())

if file_overseas and file_analiza:
    @st.cache_data
    def load_data(f_overseas, f_analiza):
        df_ov = pd.read_excel(f_overseas)
        df_an = pd.read_excel(f_analiza, header=1)
        return df_ov, df_an

    df_ov, df_an = load_data(file_overseas, file_analiza)

    # 1. Ekstrakcija broja otpremnice iz reference
    df_ov['Extracted_OT'] = df_ov['Reference 1'].apply(extract_ot_number)
    if 'Broj' in df_an.columns:
        df_an['Broj_int'] = pd.to_numeric(df_an['Broj'], errors='coerce')
    else:
        st.error("U Analizi izlaza nije pronađen stupac 'Broj'.")
        st.stop()

    # Inicijalizacija stupaca
    df_ov['Matched_Method'] = None
    df_ov['Matched_Broj_Otpremnice'] = None
    df_ov['Matched_Naziv_Korisnika'] = None

    # KORAK 1: Spajanje po broju otpremnice iz Reference 1
    for idx, row in df_ov.iterrows():
        ot = row['Extracted_OT']
        if pd.notna(ot):
            match = df_an[df_an['Broj_int'] == ot]
            if not match.empty:
                df_ov.loc[idx, 'Matched_Method'] = 'Broj Otpremnice'
                df_ov.loc[idx, 'Matched_Broj_Otpremnice'] = match.iloc[0]['Broj']
                df_ov.loc[idx, 'Matched_Naziv_Korisnika'] = match.iloc[0].get('Naziv korisnika', '')

    # KORAK 2: Spajanje preostalih po preciznom nazivu kupca
    unmatched_mask = df_ov['Matched_Method'].isna()
    
    df_an['Clean_Korisnik'] = df_an['Naziv korisnika'].apply(clean_name)
    df_an['Clean_Partner'] = df_an['Naziv partnera'].apply(clean_name)
    
    clean_korisnici_list = [str(x) for x in df_an['Clean_Korisnik'].dropna().unique() if len(str(x)) > 3]

    for idx, row in df_ov[unmatched_mask].iterrows():
        ov_name_raw = row['Consignee Name']
        ov_clean = clean_name(ov_name_raw)
        
        if len(ov_clean) < 3:
            continue

        matched_broj = None
        matched_naziv = None

        # A) Tražimo strogo poklapanje (ako naziv iz analize sadrži Overseas naziv ili obratno)
        for _, an_row in df_an.iterrows():
            ck = str(an_row['Clean_Korisnik'])
            cp = str(an_row['Clean_Partner'])
            if (len(ck) > 3 and (ck in ov_clean or ov_clean in ck)) or \
               (len(cp) > 3 and (cp in ov_clean or ov_clean in cp)):
                matched_broj = an_row['Broj']
                matched_naziv = an_row.get('Naziv korisnika', '')
                break

        # B) Ako nema direktnog sadrži-u-sebi, tražimo visoku sličnost (cutoff povisujemo na 0.82 da nema grešaka)
        if matched_broj is None and clean_korisnici_list:
            close_matches = difflib.get_close_matches(ov_clean, clean_korisnici_list, n=1, cutoff=0.82)
            if close_matches:
                best_match_str = close_matches[0]
                matched_row = df_an[df_an['Clean_Korisnik'] == best_match_str].iloc[0]
                matched_broj = matched_row['Broj']
                matched_naziv = matched_row.get('Naziv korisnika', '')

        if matched_broj is not None:
            df_ov.loc[idx, 'Matched_Method'] = 'Naziv Kupca'
            df_ov.loc[idx, 'Matched_Broj_Otpremnice'] = matched_broj
            df_ov.loc[idx, 'Matched_Naziv_Korisnika'] = matched_naziv

    # Definiranje konačnih skupova
    matched = df_ov[df_ov['Matched_Method'].notna()]
    unmatched = df_ov[df_ov['Matched_Method'].isna()]

    # Prikaz metrika
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ukupno stavki na fakturi", len(df_ov))
    col2.metric("Uspješno povezano", len(matched), delta=f"{len(matched)/len(df_ov)*100:.1f}%")
    col3.metric("Nije uspjelo povezati", len(unmatched), delta=f"-{len(unmatched)/len(df_ov)*100:.1f}%", delta_color="inverse")
    
    total_cost_unmatched = unmatched['Total Price'].sum() if 'Total Price' in unmatched.columns else 0
    col4.metric("Iznos nepovezanih troškova", f"{total_cost_unmatched:,.2f} EUR")

    # Tabs za pregled
    tab1, tab2, tab3 = st.tabs(["❌ Još uvijek nepovezane", "✅ Uspješno povezano", "📊 Detalji spajanja"])

    with tab1:
        st.subheader("Stavke koje se nisu mogle povezati ni po broju ni po nazivu kupca")
        if len(unmatched) > 0:
            display_cols_ov = [c for c in ['Invoice Number', 'Shipping Date', 'Consignee Name', 'Consignee Town', 'Reference 1', 'Total Price'] if c in unmatched.columns]
            st.dataframe(unmatched[display_cols_ov].astype(str), use_container_width=True)

            csv = unmatched.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Preuzmi nepovezane stavke (CSV)",
                data=csv,
                file_name="nepovezane_overseas.csv",
                mime="text/csv",
            )
        else:
            st.success("Sve stavke su uspješno povezane!")

    with tab2:
        st.subheader("Uspješno povezane stavke")
        if len(matched) > 0:
            display_cols_matched = [c for c in ['Invoice Number', 'Matched_Method', 'Matched_Broj_Otpremnice', 'Consignee Name', 'Matched_Naziv_Korisnika', 'Total Price'] if c in matched.columns]
            st.dataframe(matched[display_cols_matched].astype(str), use_container_width=True)
        else:
            st.info("Nema povezanih stavki.")

    with tab3:
        st.subheader("Sve stavke s označenom metodom spajanja")
        summary_cols = [c for c in ['Invoice Number', 'Reference 1', 'Consignee Name', 'Matched_Method', 'Matched_Broj_Otpremnice'] if c in df_ov.columns]
        st.dataframe(df_ov[summary_cols].astype(str), use_container_width=True)

else:
   st.info("👈 Molimo vas da učitate obje Excel datoteke u bočnoj traci (sidebaru).")
