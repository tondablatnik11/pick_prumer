import streamlit as st
import pandas as pd

# Nastavení vzhledu stránky
st.set_page_config(page_title="Analýza Pickování", layout="wide")

def main():
    st.title("📦 Analýza Pickování")
    st.write("Tato aplikace vyfiltruje zakázky, které obsahují **pouze jeden materiál** a byly pickovány **na paletu** (ověřeno dle formátu Certificate Number). Následně vypočítá průměry.")

    # Komponenta pro nahrání souboru přes web (odstraní chybu FileNotFoundError)
    uploaded_file = st.file_uploader("Nahrajte exportovaný soubor (CSV nebo Excel)", type=['csv', 'xlsx'])

    if uploaded_file is not None:
        try:
            # 1. Načtení dat na základě typu souboru
            with st.spinner('Načítám a zpracovávám data...'):
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df = pd.read_excel(uploaded_file, dtype=str)
                
                # Očištění dat od prázdných zakázek
                df = df.dropna(subset=['Delivery']).copy()
                
                # Převedení množství na čísla pro výpočty
                df['Qty'] = pd.to_numeric(df['Act.qty (dest)'], errors='coerce').fillna(0)

                # 2. Funkce pro ověření certifikátu
                def is_valid_cert(certs):
                    valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
                    # Nesmí být prázdné
                    if len(valid_certs) == 0:
                        return False
                    # Žádný z certifikátů nesmí začínat na '460'
                    for c in valid_certs:
                        if c.startswith('460'):
                            return False
                    return True

                # 3. Seskupení dat podle zakázky (Delivery)
                grouped = df.groupby('Delivery').agg(
                    num_materials=('Material', 'nunique'),
                    certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
                    total_qty=('Qty', 'sum'),
                    num_positions=('Source Storage Bin', 'nunique')
                )

                # 4. Filtrace zakázek (1 materiál a platný certifikát)
                filtered_orders = grouped[
                    (grouped['num_materials'] == 1) & 
                    (grouped['certs'].apply(is_valid_cert))
                ]

                total_filtered_orders = len(filtered_orders)

                # 5. Zobrazení výsledků
                st.divider()
                if total_filtered_orders > 0:
                    avg_qty = filtered_orders['total_qty'].mean()
                    avg_pos = filtered_orders['num_positions'].mean()

                    st.subheader("📊 Výsledky analýzy")
                    
                    # Zobrazení metrik vedle sebe
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Počet vyfiltrovaných zakázek", f"{total_filtered_orders:,}".replace(',', ' '))
                    col2.metric("Průměrný počet kusů na zakázku", f"{avg_qty:.2f}")
                    col3.metric("Průměrný počet pozic na zakázku", f"{avg_pos:.2f}")

                else:
                    st.warning("Nenalezeny žádné zakázky odpovídající zadaným kritériím.")

        except Exception as e:
            st.error(f"Došlo k chybě při zpracování souboru. Zkontrolujte formát dat. Detail chyby: {e}")

if __name__ == "__main__":
    main()
