import streamlit as st
import pandas as pd

# Nastavení vzhledu stránky
st.set_page_config(page_title="Analýza Pickování", layout="wide")

def main():
    st.title("📦 Analýza Pickování")
    st.write("Tato aplikace vyfiltruje zakázky, které obsahují **pouze jeden materiál** a byly pickovány **na paletu** (ověřeno dle formátu Certificate Number). Zobrazuje průměry a také statistiku TOP 20 materiálů z celého souboru.")

    # Komponenta pro nahrání souboru přes web
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

                # ==========================================
                # SEKCE 1: FILTRACE ZAKÁZEK (1 MATERIÁL)
                # ==========================================
                
                # Funkce pro ověření certifikátu
                def is_valid_cert(certs):
                    valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
                    if len(valid_certs) == 0:
                        return False
                    for c in valid_certs:
                        if c.startswith('460'):
                            return False
                    return True

                # Seskupení dat podle zakázky (Delivery)
                grouped = df.groupby('Delivery').agg(
                    num_materials=('Material', 'nunique'),
                    material=('Material', 'first'),
                    certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
                    total_qty=('Qty', 'sum'),
                    num_positions=('Source Storage Bin', 'nunique')
                )

                # Filtrace zakázek (1 materiál a platný certifikát)
                filtered_orders = grouped[
                    (grouped['num_materials'] == 1) & 
                    (grouped['certs'].apply(is_valid_cert))
                ]

                total_filtered_orders = len(filtered_orders)

                st.divider()
                st.subheader("🎯 Analýza paletových zakázek (1 materiál)")
                
                if total_filtered_orders > 0:
                    avg_qty = filtered_orders['total_qty'].mean()
                    avg_pos = filtered_orders['num_positions'].mean()
                    
                    # Zobrazení metrik vedle sebe
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Počet vyfiltrovaných zakázek", f"{total_filtered_orders:,}".replace(',', ' '))
                    col2.metric("Průměrný počet kusů na zakázku", f"{avg_qty:.2f}")
                    col3.metric("Průměrný počet pozic na zakázku", f"{avg_pos:.2f}")

                    # Zobrazení detailů s materiálem
                    with st.expander("Zobrazit detail vyfiltrovaných zakázek (včetně materiálu)"):
                        display_df = filtered_orders[['material', 'certs', 'total_qty', 'num_positions']].copy()
                        display_df.rename(columns={
                            'material': 'Materiál',
                            'certs': 'Certifikáty',
                            'total_qty': 'Celkem kusů',
                            'num_positions': 'Počet pozic'
                        }, inplace=True)
                        st.dataframe(display_df, use_container_width=True)
                else:
                    st.warning("Nenalezeny žádné zakázky odpovídající zadaným kritériím pro 1 materiál na paletu.")

                # ==========================================
                # SEKCE 2: TOP 20 NEJPICKOVANĚJŠÍCH MATERIÁLŮ
                # ==========================================
                st.divider()
                st.subheader("🏆 TOP 20 nejpickovanějších materiálů (ze všech dat)")
                
                # Agregace všech dat pro zjištění TOP 20
                top_materials = df.groupby('Material').agg(
                    pocet_picku=('Material', 'count'),  # Kolikrát byl na řádku (četnost pickování)
                    celkove_mnozstvi=('Qty', 'sum')     # Kolik kusů se celkem napickovalo
                ).reset_index()

                # Seřazení podle četnosti pickování
                top_20 = top_materials.sort_values(by='pocet_picku', ascending=False).head(20)

                # Přejmenování sloupců pro hezčí výpis
                top_20.rename(columns={
                    'Material': 'Materiál',
                    'pocet_picku': 'Počet picknutí (operací)',
                    'celkove_mnozstvi': 'Celkové množství (ks)'
                }, inplace=True)

                col_top1, col_top2 = st.columns([1, 1])

                with col_top1:
                    st.write("**Tabulka TOP 20**")
                    # Skryjeme index (číslování řádků od nuly) pro čistší vzhled
                    st.dataframe(top_20, use_container_width=True, hide_index=True)

                with col_top2:
                    st.write("**Graf četnosti pickování**")
                    # Graf chceme primárně podle "Počtu picknutí", nastavíme Materiál jako osu X (index)
                    st.bar_chart(top_20.set_index('Materiál')['Počet picknutí (operací)'])

        except Exception as e:
            st.error(f"Došlo k chybě při zpracování souboru. Zkontrolujte formát dat. Detail chyby: {e}")

if __name__ == "__main__":
    main()
