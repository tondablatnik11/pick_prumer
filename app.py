import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Chytrá Analýza Pickování", layout="wide")

def main():
    st.title("📦 Chytrá Analýza Pickování & Zátěže")
    st.write("Nahrajte **Pick report** i **MARM report**. Aplikace z výpočtů **zcela vyřadí zakázky**, kde byla pickována celá jednotka (Removal of total SU = X). U zbytku chytře využívá váhy a balení pro výpočet hmatů.")

    uploaded_files = st.file_uploader("Nahrajte soubory (Pick report a MARM report)", type=['csv', 'xlsx'], accept_multiple_files=True)

    if uploaded_files and len(uploaded_files) > 0:
        df_pick = None
        df_marm = None

        with st.spinner('Načítám a identifikuji soubory...'):
            for file in uploaded_files:
                if file.name.lower().endswith('.csv'):
                    temp_df = pd.read_csv(file, dtype=str)
                else:
                    temp_df = pd.read_excel(file, dtype=str)
                
                if 'Delivery' in temp_df.columns:
                    df_pick = temp_df
                elif 'Numerator' in temp_df.columns and 'Alternative Unit of Measure' in temp_df.columns:
                    df_marm = temp_df

        if df_pick is None:
            st.error("Nepodařilo se najít Pick report (chybí sloupec 'Delivery').")
            return
            
        if df_marm is None:
            st.warning("Nebyl nahrán MARM report. Výpočet pohybů bude pouze orientační bez krabic a vah.")

        # Očištění dat
        df_pick = df_pick.dropna(subset=['Delivery', 'Material']).copy()
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)

        # ==========================================
        # NOVÉ: VYŘAZENÍ ZAKÁZEK S 'X'
        # ==========================================
        # Najdeme všechna čísla zakázek, které mají alespoň na jednom řádku 'X'
        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().str.upper()
        zakazky_s_x = df_pick[df_pick['Removal of total SU'] == 'X']['Delivery'].unique()
        
        # Odstraníme tyto zakázky kompletně z našeho datasetu
        df_pick = df_pick[~df_pick['Delivery'].isin(zakazky_s_x)].copy()
        
        st.info(f"Z dat bylo kompletně vyloučeno **{len(zakazky_s_x)} zakázek**, protože obsahovaly odběr celé jednotky ('X').")

        # ==========================================
        # ZPRACOVÁNÍ MARM DAT A NAPOJENÍ
        # ==========================================
        box_dict = {}
        weight_dict = {}

        if df_marm is not None:
            df_boxes = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK', 'KAR', 'KART', 'PAK', 'VPE', 'CAR', 'BLO'])]
            df_boxes['Numerator'] = pd.to_numeric(df_boxes['Numerator'], errors='coerce').fillna(0)
            
            box_sizes = df_boxes.groupby('Material')['Numerator'].max().to_dict()
            box_dict = {mat: int(size) for mat, size in box_sizes.items() if size > 1}

            df_st = df_marm[df_marm['Alternative Unit of Measure'].isin(['ST', 'PCE', 'KS'])]
            df_st['Gross Weight'] = pd.to_numeric(df_st['Gross Weight'], errors='coerce').fillna(0)
            
            def to_kg(row):
                w = row['Gross Weight']
                u = str(row['Unit of Weight']).upper()
                if u == 'G': return w / 1000.0
                if u == 'MG': return w / 1000000.0
                return w

            df_st['Weight_KG'] = df_st.apply(to_kg, axis=1)
            weight_dict = df_st.groupby('Material')['Weight_KG'].first().to_dict()

        df_pick['Box_Size'] = df_pick['Material'].map(box_dict).fillna(0)
        df_pick['Piece_Weight_KG'] = df_pick['Material'].map(weight_dict).fillna(0)

        # ==========================================
        # POSTRANNÍ PANEL: NASTAVENÍ A FILTRY
        # ==========================================
        st.sidebar.header("⚙️ Fyzické limity pickera")
        
        limit_vahy = st.sidebar.number_input(
            "Od jaké váhy musí brát kusy po 1? (kg)", 
            min_value=0.1, max_value=20.0, value=2.0, step=0.5,
            help="Pokud 1 volný kus váží více než tato hodnota, nelze jich vzít více do jedné ruky."
        )

        kusy_na_hmat = st.sidebar.slider(
            "Max kusů do ruky (pro lehké díly)", 
            min_value=1, max_value=20, value=3, step=1,
            help="Kolik drobných kusů dokáže picker chytit do hrsti najednou?"
        )

        st.sidebar.divider()
        unique_materials = sorted(df_pick['Material'].unique().tolist())
        excluded_materials = st.sidebar.multiselect(
            "Vyloučit materiály z výpočtů:",
            options=unique_materials,
            default=[]
        )
        
        if excluded_materials:
            df_pick = df_pick[~df_pick['Material'].isin(excluded_materials)]

        # ==========================================
        # CHYTRÝ VÝPOČET POHYBŮ (Bez 'X', to už je vyřazeno)
        # ==========================================
        def spocitej_pohyby(row):
            qty = row['Qty']
            if qty <= 0:
                return 0
            
            pohyby = 0
            zbytek = qty
            box_size = row['Box_Size']
            
            # 1. Zpracování po celých kartonech (menší pod-krabice uvnitř pozice)
            if box_size > 1 and zbytek >= box_size:
                plne_kartony = zbytek // box_size
                pohyby += plne_kartony
                zbytek = zbytek % box_size
                
            # 2. Zpracování zbylých volných kusů (podle váhy)
            if zbytek > 0:
                vaha_kusu = row['Piece_Weight_KG']
                if vaha_kusu >= limit_vahy:
                    pohyby += zbytek # Těžké bere po 1 ks
                else:
                    pohyby += np.ceil(zbytek / kusy_na_hmat) # Lehké bere po hrstech
                    
            return pohyby

        df_pick['Pohyby_Rukou'] = df_pick.apply(spocitej_pohyby, axis=1)
        df_pick['Celkova_Vaha_KG'] = df_pick['Qty'] * df_pick['Piece_Weight_KG']

        # ==========================================
        # SEKCE 1: FILTRACE ZAKÁZEK (1 MATERIÁL)
        # ==========================================
        def is_valid_cert(certs):
            valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
            if len(valid_certs) == 0: return False
            for c in valid_certs:
                if c.startswith('460'): return False
            return True

        # Seskupení
        grouped = df_pick.groupby('Delivery').agg(
            num_materials=('Material', 'nunique'),
            material=('Material', 'first'),
            certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
            total_qty=('Qty', 'sum'),
            num_positions=('Source Storage Bin', 'nunique'),
            celkem_pohybu=('Pohyby_Rukou', 'sum'),
            vaha_zakazky=('Celkova_Vaha_KG', 'sum')
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
            avg_pohybu = filtered_orders['celkem_pohybu'].mean()
            avg_vaha = filtered_orders['vaha_zakazky'].mean()
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Počet vyfiltrovaných zakázek", f"{total_filtered_orders:,}".replace(',', ' '))
            col2.metric("Průměrně kusů / zakázka", f"{avg_qty:.1f} ks")
            col3.metric("Průměrně pozic / zakázka", f"{avg_pos:.2f}")
            col4.metric("Průměrně pohybů / zakázka", f"{avg_pohybu:.1f} hmatů")

            with st.expander("Zobrazit detail vyfiltrovaných zakázek"):
                display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'vaha_zakazky', 'certs']].copy()
                display_df.rename(columns={
                    'material': 'Materiál',
                    'total_qty': 'Celkem kusů',
                    'celkem_pohybu': 'Pohyby rukou',
                    'vaha_zakazky': 'Odhad váhy (kg)',
                    'certs': 'Certifikáty'
                }, inplace=True)
                st.dataframe(display_df, use_container_width=True)
        else:
            st.warning("Nenalezeny žádné zakázky odpovídající zadaným kritériím.")

        # ==========================================
        # SEKCE 2: TOP 50 NEJNÁROČNĚJŠÍCH MATERIÁLŮ
        # ==========================================
        st.divider()
        st.subheader("🏆 TOP 50 fyzicky nejnáročnějších materiálů (ze všech volných picků)")
        
        if not df_pick.empty:
            top_materials = df_pick.groupby('Material').agg(
                pocet_picku=('Material', 'count'),
                celkove_mnozstvi=('Qty', 'sum'),
                celkem_pohybu=('Pohyby_Rukou', 'sum'),
                celkova_natacena_vaha=('Celkova_Vaha_KG', 'sum')
            ).reset_index()

            top_50 = top_materials.sort_values(by='celkem_pohybu', ascending=False).head(50)

            top_50.rename(columns={
                'Material': 'Materiál',
                'pocet_picku': 'Příjezdy na pozici (řádky)',
                'celkem_pohybu': 'Fyzické pohyby rukou',
                'celkove_mnozstvi': 'Kusů celkem',
                'celkova_natacena_vaha': 'Zvednuto (kg)'
            }, inplace=True)

            col_top1, col_top2 = st.columns([1.5, 1])

            with col_top1:
                st.dataframe(top_50.style.format({
                    "Zvednuto (kg)": "{:.1f}",
                    "Fyzické pohyby rukou": "{:.0f}"
                }), use_container_width=True, hide_index=True)

            with col_top2:
                st.bar_chart(top_50.set_index('Materiál')['Fyzické pohyby rukou'])

if __name__ == "__main__":
    main()
