import streamlit as st
import pandas as pd
import numpy as np

# Nastavení vzhledu stránky
st.set_page_config(page_title="Analýza Pickování", layout="wide")

def main():
    st.title("📦 Analýza Pickování & Fyzické pohyby")
    st.write("Aplikace filtruje zakázky (1 materiál, na paletu), vypočítá průměry, TOP materiály a odhaduje **počet fyzických pohybů pickera**.")

    uploaded_file = st.file_uploader("Nahrajte exportovaný soubor (CSV nebo Excel)", type=['csv', 'xlsx'])

    if uploaded_file is not None:
        try:
            with st.spinner('Načítám a zpracovávám data...'):
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df = pd.read_excel(uploaded_file, dtype=str)
                
                df = df.dropna(subset=['Delivery']).copy()
                df = df.dropna(subset=['Material'])
                df['Qty'] = pd.to_numeric(df['Act.qty (dest)'], errors='coerce').fillna(0)

            # ==========================================
            # POSTRANNÍ PANEL: NASTAVENÍ A FILTRY
            # ==========================================
            st.sidebar.header("⚙️ Nastavení výpočtu pohybů")
            st.sidebar.write("Jak se počítají fyzické pohyby rukou:")
            st.sidebar.markdown("- **Celá krabice** (`Removal of total SU`='X') = 1 pohyb")
            st.sidebar.markdown("- **Jednotlivé kusy** = Počet kusů / Počet ks na jeden hmat")
            
            # Posuvník: Kolik kusů zvládne picker vzít najednou do ruky?
            kusy_na_hmat = st.sidebar.slider(
                "Kolik kusů průměrně vezme picker do ruky (hmat)?", 
                min_value=1, max_value=10, value=1, step=1
            )

            st.sidebar.divider()
            
            # Filtr pro vyloučení materiálů
            unique_materials = sorted(df['Material'].unique().tolist())
            excluded_materials = st.sidebar.multiselect(
                "Vyberte materiály k vyloučení ze všech výpočtů:",
                options=unique_materials,
                default=[]
            )
            
            if excluded_materials:
                df = df[~df['Material'].isin(excluded_materials)]

            # ==========================================
            # VÝPOČET FYZICKÝCH POHYBŮ PRO KAŽDÝ ŘÁDEK
            # ==========================================
            # Pokud vybírá celou jednotku ('X'), je to 1 pohyb. 
            # Jinak vezmeme počet kusů, vydělíme počtem ks na hmat a zaokrouhlíme nahoru.
            df['Pohyby_Rukou'] = np.where(
                df['Removal of total SU'] == 'X', 
                1, 
                np.ceil(df['Qty'] / kusy_na_hmat)
            )

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
                num_positions=('Source Storage Bin', 'nunique'),
                celkem_pohybu=('Pohyby_Rukou', 'sum') # Přidán součet pohybů
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
                avg_pohybu = filtered_orders['celkem_pohybu'].mean() # Průměrný počet pohybů
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Počet vyfiltrovaných zakázek", f"{total_filtered_orders:,}".replace(',', ' '))
                col2.metric("Průměrný počet kusů na zakázku", f"{avg_qty:.1f}")
                col3.metric("Průměrný počet pozic na zakázku", f"{avg_pos:.2f}")
                col4.metric("Průměrně fyzických pohybů na zakázku", f"{avg_pohybu:.1f}") # Nová metrika

                with st.expander("Zobrazit detail vyfiltrovaných zakázek (včetně pohybů)"):
                    display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'num_positions', 'certs']].copy()
                    display_df.rename(columns={
                        'material': 'Materiál',
                        'total_qty': 'Celkem kusů',
                        'celkem_pohybu': 'Odhad pohybů',
                        'num_positions': 'Počet pozic',
                        'certs': 'Certifikáty'
                    }, inplace=True)
                    st.dataframe(display_df, use_container_width=True)
            else:
                st.warning("Nenalezeny žádné zakázky odpovídající zadaným kritériím.")

            # ==========================================
            # SEKCE 2: TOP 50 NEJPICKOVANĚJŠÍCH MATERIÁLŮ
            # ==========================================
            st.divider()
            st.subheader("🏆 TOP 50 nejnáročnějších materiálů")
            st.write("*(Do této statistiky nejsou započítány materiály vyloučené v postranním panelu)*")
            
            top_materials = df.groupby('Material').agg(
                pocet_picku=('Material', 'count'),          # kolikrát se pro materiál šlo (počet řádků)
                celkove_mnozstvi=('Qty', 'sum'),            # celkem kusů
                celkem_pohybu=('Pohyby_Rukou', 'sum')       # celkem odhadovaných pohybů rukou
            ).reset_index()

            # Změna: Teď tabulku řadíme podle toho, kolikrát musel picker pro daný materiál něco udělat rukama
            top_50 = top_materials.sort_values(by='celkem_pohybu', ascending=False).head(50)

            top_50.rename(columns={
                'Material': 'Materiál',
                'pocet_picku': 'Počet příjezdů (řádků)',
                'celkem_pohybu': 'Fyzické pohyby rukou',
                'celkove_mnozstvi': 'Celkové množství (ks)'
            }, inplace=True)

            col_top1, col_top2 = st.columns([1, 1])

            with col_top1:
                st.write("**Tabulka TOP 50 (seřazeno dle fyzické náročnosti)**")
                st.dataframe(top_50, use_container_width=True, hide_index=True)

            with col_top2:
                st.write("**Graf fyzické náročnosti pickování**")
                st.bar_chart(top_50.set_index('Materiál')['Fyzické pohyby rukou'])

        except Exception as e:
            st.error(f"Došlo k chybě při zpracování souboru. Detail: {e}")

if __name__ == "__main__":
    main()
