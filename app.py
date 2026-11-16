import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Usklađivanje Overseas faktura i Izlaza", layout="wide")

st.title("📦 Pametno usklađivanje Overseas faktura i Izlaza")
st.markdown("Aplikacija spaja po broju otpremnice, a preostale stavke detaljno uspoređuje po **Adresi**, **Mjestu** i **Nazivu korisnika**!")

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

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    for word in ['d.o.o.', 'd.d.', 'j.d.o.o.', '-', 'isporuka', 'skladište', 'doo', 'dd', ',', '.', '/']:
        text = text.replace(word, ' ')
    return ' '.join(text.split())

def clean_address(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    # Uklanjanje čestih oznaka u adresama radi bolje usporedbe
    for word in ['ulica', 'ul.', 'dr.', 'bb', 'kbr', 'k.br.', ',', '.', '/']:
        text = text.replace(word, ' ')
    return ' '.join(text.split())

if file_overseas and file_analiza:
    @st.cache_data
    def load_data(f_overseas, f_analiza):
        df_ov = pd.read_excel(f_overseas)
        df_raw = pd.read_excel(f_analiza, header=None)
        header_row = 1
        for idx, row in df_raw.iterrows():
            row_str = row.astype(str).str.cat(sep=' ').lower()
            if 'broj' in row_str and ('naziv' in row_str or 'korisnik' in row_str):
                header_row = idx
                break
        df_an = pd.read_excel(f_analiza, header=header_row)
        return df_ov, df_an

    df_ov, df_an = load_data(file_overseas, file_analiza)

    # Pronalazak stupca za broj otpremnice u Analizi izlaza
    broj_col = None
    for col in df_an.columns:
        if str(col).strip().lower() in ['broj', 'broj dokumenta', 'otpremnica']:
            broj_col = col
            break
    if not broj_col:
        broj_col = df_an.columns[1] if len(df_an.columns) > 1 else df_an.columns[0]

    def clean_broj(val):
        if pd.isna(val):
            return None
        s = str(val).replace('.', '').replace(',', '').strip()
        if s.isdigit():
            return int(s)
        return None

    df_an['Broj_int'] = df_an[broj_col].apply(clean_broj)

    # Identifikacija stupaca u Analizi izlaza (fleksibilno pretraživanje)
    korisnik_col = next((c for c in df_an.columns if any(k in str(c).lower() for k in ['korisnik', 'naziv'])), df_an.columns[6])
    adresa_col = next((c for c in df_an.columns if any(k in str(c).lower() for k in ['adresa', 'ulica'])), None)
    mjesto_col = next((c for c in df_an.columns if any(k in str(c).lower() for k in ['mjesto', 'grad', 'posta'])), df_an.columns[9])

    # Identifikacija stupaca u Overseas datoteci
    ov_name_col = next((c for c in df_ov.columns if 'consignee name' in str(c).lower() or 'naziv' in str(c).lower()), 'Consignee Name')
    ov_street_col = next((c for c in df_ov.columns if any(k in str(c).lower() for k in ['street', 'adresa', 'ulica'])), 'Consignee Street')
    ov_town_col = next((c for c in df_ov.columns if any(k in str(c).lower() for k in ['town', 'city', 'mjesto', 'grad'])), 'Consignee Town')

    # Priprema čistih podataka za usporedbu
    df_an['Clean_Korisnik'] = df_an[korisnik_col].apply(clean_text)
    df_an['Clean_Adresa'] = df_an[adresa_col].apply(clean_address) if adresa_col else ""
    df_an['Clean_Mjesto'] = df_an[mjesto_col].apply(clean_text) if mjesto_col else ""

    df_ov['Clean_Consignee'] = df_ov[ov_name_col].apply(clean_text) if ov_name_col in df_ov.columns else ""
    df_ov['Clean_Street'] = df_ov[ov_street_col].apply(clean_address) if ov_street_col in df_ov.columns else ""
    df_ov['Clean_Town'] = df_ov[ov_town_col].apply(clean_text) if ov_town_col in df_ov.columns else ""

    # Ekstrakcija broja otpremnice iz Overseas reference
    ref_col = 'Reference 1' if 'Reference 1' in df_ov.columns else df_ov.columns[0]
    df_ov['Extracted_OT'] = df_ov[ref_col].apply(extract_ot_number)

    df_ov['Matched_Method'] = None
    df_ov['Matched_Broj_Otpremnice'] = None
    df_ov['Matched_Naziv_Korisnika'] = None

    # KORAK 1: Spajanje po broju otpremnice
    for idx, row in df_ov.iterrows():
        ot = row['Extracted_OT']
        if pd.notna(ot):
            match = df_an[df_an['Broj_int'] == ot]
            if not match.empty:
                df_ov.loc[idx, 'Matched_Method'] = 'Broj Otpremnice'
                df_ov.loc[idx, 'Matched_Broj_Otpremnice'] = match.iloc[0][broj_col]
                df_ov.loc[idx, 'Matched_Naziv_Korisnika'] = match.iloc[0][korisnik_col]

    # KORAK 2: Spajanje preostalih po Adresi, Mjestu i Nazivu korisnika
    unmatched_mask = df_ov['Matched_Method'].isna()
    for idx, row in df_ov[unmatched_mask].iterrows():
        c_name = row['Clean_Consignee']
        c_street = row['Clean_Street']
        c_town = row['Clean_Town']
        
        matched_row = None
        for _, an_row in df_an.iterrows():
            ak = str(an_row['Clean_Korisnik'])
            aa = str(an_row['Clean_Adresa'])
            am = str(an_row['Clean_Mjesto'])
            
            # Provjera poklapanja: Ulica + Mjesto + Naziv
            score = 0
            if len(c_street) > 3 and aa and (aa in c_street or c_street in aa):
                score += 2  # Adresa nosi veću težinu
            if len(c_town) > 2 and am and (am in c_town or c_town in am):
                score += 1  # Mjesto
            if len(c_name) > 3 and ak and (ak in c_name or c_name in ak):
                score += 2  # Naziv

            # Uvjet za uspješno spajanje (npr. poklopljena adresa/naziv uz mjesto, ili visoka podudarnost)
            if score >= 3: 
                matched_row = an_row
                break

        if matched_row is not None:
            df_ov.loc[idx, 'Matched_Method'] = 'Adresa + Mjesto + Naziv'
            df_ov.loc[idx, 'Matched_Broj_Otpremnice'] = matched_row[broj_col]
            df_ov.loc[idx, 'Matched_Naziv_Korisnika'] = matched_row[korisnik_col]

    # Definiranje uspješnih i neuspjelih
    matched = df_ov[df_ov['Matched_Method'].notna()]
    unmatched = df_ov[df_ov['Matched_Method'].isna()]

    # Metrike
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ukupno stavki na fakturi", len(df_ov))
    col2.metric("Uspješno povezano", len(matched), delta=f"{len(matched)/len(df_ov)*100:.1f}%")
    col3.metric("Nije uspjelo povezati", len(unmatched), delta=f"-{len(unmatched)/len(df_ov)*100:.1f}%", delta_color="inverse")
    
    total_cost_matched = matched['Total Price'].sum() if 'Total Price' in matched.columns else 0
    col4.metric("Iznos povezanih troškova", f"{total_cost_matched:,.2f} EUR")

    # Kartice
    tab1, tab2, tab3 = st.tabs(["✅ Uspješno povezane stavke", "❌ Još uvijek nepovezane", "📊 Sve stavke"])

    with tab1:
        st.subheader("Uspješno povezane stavke")
        if len(matched) > 0:
            display_cols = ['Invoice Number', 'Matched_Method', 'Matched_Broj_Otpremnice', ov_name_col, 'Matched_Naziv_Korisnika', 'Total Price']
            available_cols = [c for c in display_cols if c in matched.columns]
            st.dataframe(matched[available_cols].astype(str), use_container_width=True)

            csv_match = matched.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Preuzmi povezane stavke (CSV)",
                data=csv_match,
                file_name="povezane_overseas_stavke.csv",
                mime="text/csv",
            )
        else:
            st.info("Nema povezanih stavki.")

    with tab2:
        st.subheader("Stavke koje se nisu mogle povezati")
        if len(unmatched) > 0:
            display_cols_un = [c for c in ['Invoice Number', 'Shipping Date', ov_name_col, ov_street_col, ov_town_col, ref_col, 'Total Price'] if c in unmatched.columns]
            st.dataframe(unmatched[display_cols_un].astype(str), use_container_width=True)

            csv_un = unmatched.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Preuzmi nepovezane stavke (CSV)",
                data=csv_un,
                file_name="nepovezane_overseas.csv",
                mime="text/csv",
            )
        else:
            st.success("Sve stavke s fakture su uspješno povezane!")

    with tab3:
        st.subheader("Pregled svih stavki i metoda spajanja")
        summary_cols = [c for c in ['Invoice Number', ref_col, ov_name_col, 'Matched_Method', 'Matched_Broj_Otpremnice', 'Total Price'] if c in df_ov.columns]
        st.dataframe(df_ov[summary_cols].astype(str), use_container_width=True)

else:
   st.info("👈 Molimo vas da učitate obje Excel datoteke u bočnoj traci (sidebaru).")
