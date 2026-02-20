import streamlit as st
import pandas as pd

# Nastavení vzhledu stránky (musí být hned na začátku)
st.set_page_config(page_title="Analýza Pickování", layout="wide")

def main():
    st.title("📦 Analýza Pickování")
    st.write("Tato aplikace vyfiltruje zakázky (1 materiál, na paletu), vypočítá průměry a zobrazí **TOP 50 materiálů**.")

    # Komponenta pro nahrání souboru přes web
    uploaded_file = st.file_uploader("Nahrajte exportovaný soubor (CSV nebo Excel)", type=['csv', 'xlsx'])

    if uploaded_file is not None:
        try:
            # 1. Načtení dat
            with st.spinner('Načítám a zpracovávám data...'):
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df = pd.read_excel(uploaded_file, dtype=str)
                
                # Očištění dat od prázdných zakázek a chybějících materiálů
                df = df.dropna(subset=['Delivery']).copy()
                df = df.dropna(subset=['Material'])
                
                # Převedení množství na čísla pro výpočty
                df['Qty'] = pd.to_numeric(df['Act.qty (dest)'], errors='coerce').fillna(0)

            # ==========================================
            # POSTranní panel: VYLOUČENÍ MATERIÁLŮ
            # ==========================================
            st.sidebar.header("⚙️ Nastavení filtrů")
            st.sidebar.write("Vybrané materiály budou **kompletně vyloučeny** ze všech výpočtů (z průměrů i z TOP 50).")
            
            # Získání seřazeného seznamu všech unikátních materiálů
            unique_materials = sorted(df['Material'].unique().tolist())
            
            # Multiselect s možností vyhledávání (lze zadat text nebo vybrat ze seznamu)
            excluded_materials = st.sidebar.multiselect(
                "Vyberte materiály k vyloučení:",
                options=unique_materials,
                default=[]
            )
            
            # Aplikace filtru: odstraníme z dat všechny řádky, které obsahují vyloučené materiály
            if excluded_materials:
                df = df[~df['Material'].isin(excluded_materials)]
                st.sidebar.success(f"Vyloučeno {len(excluded_materials)} materiálů z výpočtů.")

            # ==========================================
            # SEKCE 1: FILTRACE ZAKÁZEK (1 MATERIÁL)
            # ==========================================
            def is_valid_cert(certs):
                valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
                if len(valid_certs) == 0:
                    return False
                for c in valid_certs:
                    if c.startswith('460'):
                        return False
                return True

            grouped = df.groupby('Delivery').agg(
                num_materials=('Material', 'nunique'),
                material=('Material', 'first'),
                certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
                total_qty=('Qty', 'sum'),
                num_positions=('Source Storage Bin', 'nunique')
            )

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
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Počet vyfiltrovaných zakázek", f"{total_filtered_orders:,}".replace(',', ' '))
                col2.metric("Průměrný počet kusů na zakázku", f"{avg_qty:.2f}")
                col3.metric("Průměrný počet pozic na zakázku", f"{avg_pos:.2f}")

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
                st.warning("Nenalezeny žádné zakázky odpovídající zadaným kritériím.")

            # ==========================================
            # SEKCE 2: TOP 50 NEJPICKOVANĚJŠÍCH MATERIÁLŮ
            # ==========================================
            st.divider()
            st.subheader("🏆 TOP 50 nejpickovanějších materiálů")
            st.write("*(Do této statistiky nejsou započítány materiály, které jste vyloučili v postranním panelu)*")
            
            top_materials = df.groupby('Material').agg(
                pocet_picku=('Material', 'count'),
                celkove_mnozstvi=('Qty', 'sum')
            ).reset_index()

            # Změna z .head(20) na .head(50)
            top_50 = top_materials.sort_values(by='pocet_picku', ascending=False).head(50)

            top_50.rename(columns={
                'Material': 'Materiál',
                'pocet_picku': 'Počet picknutí (operací)',
                'celkove_mnozstvi': 'Celkové množství (ks)'
            }, inplace=True)

            col_top1, col_top2 = st.columns([1, 1])

            with col_top1:
                st.write("**Tabulka TOP 50**")
                st.dataframe(top_50, use_container_width=True, hide_index=True)

            with col_top2:
                st.write("**Graf četnosti pickování**")
                st.bar_chart(top_50.set_index('Materiál')['Počet picknutí (operací)'])

        except Exception as e:
            st.error(f"Došlo k chybě při zpracování souboru. Detail: {e}")

if __name__ == "__main__":
    main()
