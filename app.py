import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. Nastavení stránky na profesionální vzhled
st.set_page_config(page_title="Analýza Skladové Zátěže | Warehouse Analysis", page_icon="📈", layout="wide")

# Inicializace jazyka v session_state
if 'lang' not in st.session_state:
    st.session_state.lang = 'cs'

# Slovník překladů
TEXTS = {
    'cs': {
        'switch_lang': "🇬🇧 Switch to English",
        'title': "📈 Analýza Skladové Zátěže & Pickování",
        'desc': "Tento analytický nástroj měří **skutečnou fyzickou náročnost vychystávání** ve skladu. Spojením provozních dat (Pick report) a kmenových dat (MARM) eliminuje zkreslení dané počtem kusů a přesně modeluje, kolik fyzických pohybů (hmatů) musel picker u regálu provést.",
        'upload_title': "📁 Nahrání vstupních dat",
        'upload_help': "Nahrajte exportované soubory (Pick report a MARM report)",
        'loading': "Načítám, identifikuji a propojuji soubory...",
        'err_pick': "Nepodařilo se najít Pick report (chybí sloupec 'Delivery'). Nahrajte prosím správný soubor.",
        'warn_marm': "Nebyl nahrán MARM report. Výpočet pohybů bude pouze orientační bez krabic a vah.",
        'info_clean': "💡 **Informace o čištění dat:** Z výpočtů bylo kompletně vyloučeno **{} zakázek**, u kterých došlo k odběru celé manipulační jednotky ('X'). Analyzujeme pouze ruční vychystávání z pozic.",
        'sidebar_title': "⚙️ Fyzické limity pickera",
        'weight_label': "Od jaké váhy (kg) brát kusy po 1?",
        'weight_help': "Pokud 1 volný kus váží více než tato hodnota, nelze jich vzít více do jedné ruky.",
        'dim_label': "Od jakého rozměru (cm) brát kusy po 1?",
        'dim_help': "Pokud nejdelší strana kusu (délka/šířka/výška) přesahuje tuto hodnotu v cm, bere se po jednom.",
        'hmat_label': "Max kusů do ruky (pro lehké a malé díly)",
        'hmat_help': "Kolik drobných kusů dokáže picker chytit do hrsti nebo rychle odpočítat najednou?",
        'exclude_label': "Vyloučit materiály z výpočtů:",
        'sec1_title': "🎯 Analýza paletových zakázek (1 materiál)",
        'm_orders': "Počet zakázek",
        'm_qty': "Prům. kusů / zakázka",
        'm_pos': "Prům. pozic / zakázka",
        'm_mov': "Prům. pohybů / zakázka",
        'exp_calc_title': "ℹ️ Detailní vysvětlení: Jak se počítají pohyby?",
        'exp_calc_text': """
**Výpočet automaticky prochází tyto kroky pro každý řádek pickování:**
1. **Odběr po celých kartonech:** Systém zjistí z MARMu, kolik kusů je ve standardním balení. Pokud picker vychystává např. 120 ks a karton má 50 ks, systém započítá odběr **2 celých kartonů = 2 pohyby**.
2. **Zbylé volné kusy (Těžké/Velké):** Zbylých 20 ks se zkontroluje podle váhy a rozměrů. Pokud kus váží více než limit (kg) **NEBO** je jeho nejdelší strana delší než limit (cm), picker je musí brát po jednom = **20 pohybů**.
3. **Zbylé volné kusy (Lehké/Malé):** Pokud jde o lehké a drobné díly, systém předpokládá nabrání do hrsti (nastaveno v panelu, např. 3 ks/hmat) = zbylých 20 ks znamená cca **7 pohybů**.
*Výsledkem je součet těchto pohybů.*
        """,
        'exp_detail_title': "Zobrazit detail vyfiltrovaných zakázek",
        'col_mat': "Materiál",
        'col_qty': "Celkem kusů",
        'col_mov': "Pohyby rukou",
        'col_wgt': "Odhad váhy (kg)",
        'col_max_dim': "Max rozměr (cm)",
        'col_cert': "Certifikáty",
        'no_orders': "Nenalezeny žádné zakázky odpovídající zadaným kritériím.",
        'sec2_title': "🏆 TOP 100 fyzicky nejnáročnějších materiálů (ze všech volných picků)",
        'col_lines': "Příjezdy (řádky)",
        'col_box': "Ks v balení",
        'val_loose': "Volné",
        'btn_download': "📥 Stáhnout tabulku jako Excel (.xlsx)"
    },
    'en': {
        'switch_lang': "🇨🇿 Přepnout do češtiny",
        'title': "📈 Warehouse Workload & Picking Analysis",
        'desc': "This analytical tool measures the **true physical demand of picking**. By merging operational data (Pick report) with master data (MARM), it eliminates the bias of raw piece counts and accurately models how many physical hand movements the picker had to perform at the bin.",
        'upload_title': "📁 Upload Input Data",
        'upload_help': "Upload exported files (Pick report and MARM report)",
        'loading': "Loading, identifying, and merging files...",
        'err_pick': "Could not find Pick report (missing 'Delivery' column). Please upload the correct file.",
        'warn_marm': "MARM report not uploaded. Movement calculation will be approximate without box sizes and weights.",
        'info_clean': "💡 **Data Cleaning Info:** Excluded **{} orders** completely because they contained full handling unit picks ('X'). We are analyzing only manual bin picking.",
        'sidebar_title': "⚙️ Picker's Physical Limits",
        'weight_label': "Weight limit for 1-by-1 pick (kg)",
        'weight_help': "If 1 loose piece weighs more than this, it cannot be grabbed in multiples.",
        'dim_label': "Dimension limit for 1-by-1 pick (cm)",
        'dim_help': "If the longest side (L/W/H) exceeds this value in cm, it must be picked one by one.",
        'hmat_label': "Max pieces per grab (light & small parts)",
        'hmat_help': "How many tiny pieces can the picker grab in one handful or quickly count at once?",
        'exclude_label': "Exclude materials from calculations:",
        'sec1_title': "🎯 Pallet Orders Analysis (1 material)",
        'm_orders': "Number of orders",
        'm_qty': "Avg pieces / order",
        'm_pos': "Avg bins / order",
        'm_mov': "Avg movements / order",
        'exp_calc_title': "ℹ️ Detailed Explanation: How are movements calculated?",
        'exp_calc_text': """
**The calculation automatically follows these steps for each picking line:**
1. **Full Carton Picks:** The system checks MARM for the standard box quantity. E.g., if the picker needs 120 pcs and the box holds 50 pcs, the system counts **2 full cartons = 2 movements**.
2. **Remaining Loose Pieces (Heavy/Large):** The remaining 20 pcs are checked for weight and dimensions. If a piece exceeds the weight limit (kg) **OR** its longest side exceeds the dimension limit (cm), they are picked one by one = **20 movements**.
3. **Remaining Loose Pieces (Light/Small):** If the parts are light and small, the system assumes grab-picking (set in sidebar, e.g., 3 pcs/grab) = remaining 20 pcs equal approx **7 movements**.
*The final value is the sum of these movements.*
        """,
        'exp_detail_title': "Show details of filtered orders",
        'col_mat': "Material",
        'col_qty': "Total Pieces",
        'col_mov': "Hand Movements",
        'col_wgt': "Est. Weight (kg)",
        'col_max_dim': "Max Dim. (cm)",
        'col_cert': "Certificates",
        'no_orders': "No orders found matching the criteria.",
        'sec2_title': "🏆 TOP 100 Physically Most Demanding Materials (from all loose picks)",
        'col_lines': "Lines (Visits)",
        'col_box': "Pcs in Box",
        'val_loose': "Loose",
        'btn_download': "📥 Download table as Excel (.xlsx)"
    }
}

def t(key):
    return TEXTS[st.session_state.lang][key]

def main():
    col_spacer, col_lang = st.columns([8, 1])
    with col_lang:
        if st.button(t('switch_lang')):
            st.session_state.lang = 'en' if st.session_state.lang == 'cs' else 'cs'
            st.rerun()

    st.title(t('title'))
    st.markdown(t('desc'))
    st.divider()

    with st.container():
        st.subheader(t('upload_title'))
        uploaded_files = st.file_uploader(
            t('upload_help'), 
            type=['csv', 'xlsx'], 
            accept_multiple_files=True
        )

    if uploaded_files and len(uploaded_files) > 0:
        df_pick = None
        df_marm = None

        with st.spinner(t('loading')):
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
            st.error(t('err_pick'))
            return
            
        if df_marm is None:
            st.warning(t('warn_marm'))

        df_pick = df_pick.dropna(subset=['Delivery', 'Material']).copy()
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)

        # VYŘAZENÍ ZAKÁZEK S 'X'
        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().str.upper()
        zakazky_s_x = df_pick[df_pick['Removal of total SU'] == 'X']['Delivery'].unique()
        df_pick = df_pick[~df_pick['Delivery'].isin(zakazky_s_x)].copy()
        
        st.info(t('info_clean').format(len(zakazky_s_x)))

        # ZPRACOVÁNÍ MARM DAT
        box_dict = {}
        weight_dict = {}
        dim_dict = {}

        if df_marm is not None:
            # 1. Kartony
            df_boxes = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK', 'KAR', 'KART', 'PAK', 'VPE', 'CAR', 'BLO'])]
            df_boxes['Numerator'] = pd.to_numeric(df_boxes['Numerator'], errors='coerce').fillna(0)
            box_sizes = df_boxes.groupby('Material')['Numerator'].max().to_dict()
            box_dict = {mat: int(size) for mat, size in box_sizes.items() if size > 1}

            # 2. Kusy (Váhy a Rozměry)
            df_st = df_marm[df_marm['Alternative Unit of Measure'].isin(['ST', 'PCE', 'KS'])].copy()
            
            df_st['Gross Weight'] = pd.to_numeric(df_st['Gross Weight'], errors='coerce').fillna(0)
            def to_kg(row):
                w = row['Gross Weight']
                u = str(row['Unit of Weight']).upper()
                if u == 'G': return w / 1000.0
                if u == 'MG': return w / 1000000.0
                return w
            df_st['Weight_KG'] = df_st.apply(to_kg, axis=1)
            weight_dict = df_st.groupby('Material')['Weight_KG'].first().to_dict()

            def to_cm(val, unit):
                try:
                    v = float(val)
                    u = str(unit).upper().strip()
                    if u == 'MM': return v / 10.0
                    if u == 'M': return v * 100.0
                    return v 
                except:
                    return 0.0

            df_st['L'] = df_st.apply(lambda r: to_cm(r['Length'], r['Unit of Dimension']), axis=1)
            df_st['W'] = df_st.apply(lambda r: to_cm(r['Width'], r['Unit of Dimension']), axis=1)
            df_st['H'] = df_st.apply(lambda r: to_cm(r['Height'], r['Unit of Dimension']), axis=1)
            df_st['Max_Dim_CM'] = df_st[['L', 'W', 'H']].max(axis=1) 
            dim_dict = df_st.groupby('Material')['Max_Dim_CM'].first().to_dict()

        df_pick['Box_Size'] = df_pick['Material'].map(box_dict).fillna(0)
        df_pick['Piece_Weight_KG'] = df_pick['Material'].map(weight_dict).fillna(0)
        df_pick['Piece_Max_Dim_CM'] = df_pick['Material'].map(dim_dict).fillna(0)

        # POSTRANNÍ PANEL
        st.sidebar.header(t('sidebar_title'))
        
        limit_vahy = st.sidebar.number_input(
            t('weight_label'), 
            min_value=0.1, max_value=20.0, value=2.0, step=0.5,
            help=t('weight_help')
        )
        
        limit_rozmeru = st.sidebar.number_input(
            t('dim_label'),
            min_value=1.0, max_value=200.0, value=15.0, step=1.0,
            help=t('dim_help')
        )

        kusy_na_hmat = st.sidebar.slider(
            t('hmat_label'), 
            min_value=1, max_value=20, value=3, step=1,
            help=t('hmat_help')
        )
        
        st.sidebar.divider()
        unique_materials = sorted(df_pick['Material'].unique().tolist())
        excluded_materials = st.sidebar.multiselect(
            t('exclude_label'),
            options=unique_materials,
            default=[]
        )
        
        if excluded_materials:
            df_pick = df_pick[~df_pick['Material'].isin(excluded_materials)]

        # VÝPOČET POHYBŮ S ROZMĚRY A VÁHOU
        def spocitej_pohyby(row):
            qty = row['Qty']
            if qty <= 0: return 0
            
            pohyby = 0
            zbytek = qty
            box_size = row['Box_Size']
            
            if box_size > 1 and zbytek >= box_size:
                plne_kartony = zbytek // box_size
                pohyby += plne_kartony
                zbytek = zbytek % box_size
                
            if zbytek > 0:
                vaha_kusu = row['Piece_Weight_KG']
                nejdelsi_strana = row['Piece_Max_Dim_CM']
                
                if vaha_kusu >= limit_vahy or nejdelsi_strana >= limit_rozmeru:
                    pohyby += zbytek
                else:
                    pohyby += np.ceil(zbytek / kusy_na_hmat)
                    
            return pohyby

        df_pick['Pohyby_Rukou'] = df_pick.apply(spocitej_pohyby, axis=1)
        df_pick['Celkova_Vaha_KG'] = df_pick['Qty'] * df_pick['Piece_Weight_KG']

        # SEKCE 1: FILTRACE ZAKÁZEK
        def is_valid_cert(certs):
            valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
            if len(valid_certs) == 0: return False
            for c in valid_certs:
                if c.startswith('460'): return False
            return True

        grouped = df_pick.groupby('Delivery').agg(
            num_materials=('Material', 'nunique'),
            material=('Material', 'first'),
            certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
            total_qty=('Qty', 'sum'),
            num_positions=('Source Storage Bin', 'nunique'),
            celkem_pohybu=('Pohyby_Rukou', 'sum'),
            vaha_zakazky=('Celkova_Vaha_KG', 'sum'),
            max_rozmer=('Piece_Max_Dim_CM', 'first')
        )

        filtered_orders = grouped[
            (grouped['num_materials'] == 1) & 
            (grouped['certs'].apply(is_valid_cert))
        ]

        total_filtered_orders = len(filtered_orders)

        st.divider()
        st.subheader(t('sec1_title'))
        
        if total_filtered_orders > 0:
            avg_qty = filtered_orders['total_qty'].mean()
            avg_pos = filtered_orders['num_positions'].mean()
            avg_pohybu = filtered_orders['celkem_pohybu'].mean()
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric(t('m_orders'), f"{total_filtered_orders:,}".replace(',', ' '))
            col2.metric(t('m_qty'), f"{avg_qty:.1f}")
            col3.metric(t('m_pos'), f"{avg_pos:.2f}")
            col4.metric(t('m_mov'), f"{avg_pohybu:.1f}")

            with st.expander(t('exp_calc_title')):
                st.markdown(t('exp_calc_text'))

            with st.expander(t('exp_detail_title')):
                display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'vaha_zakazky', 'max_rozmer', 'certs']].copy()
                display_df.rename(columns={
                    'material': t('col_mat'),
                    'total_qty': t('col_qty'),
                    'celkem_pohybu': t('col_mov'),
                    'vaha_zakazky': t('col_wgt'),
                    'max_rozmer': t('col_max_dim'),
                    'certs': t('col_cert')
                }, inplace=True)
                st.dataframe(display_df, use_container_width=True)
        else:
            st.warning(t('no_orders'))

        # SEKCE 2: TOP 100 NEJNÁROČNĚJŠÍCH MATERIÁLŮ
        st.divider()
        st.subheader(t('sec2_title'))
        
        if not df_pick.empty:
            top_materials = df_pick.groupby('Material').agg(
                pocet_picku=('Material', 'count'),
                celkove_mnozstvi=('Qty', 'sum'),
                celkem_pohybu=('Pohyby_Rukou', 'sum'),
                celkova_natacena_vaha=('Celkova_Vaha_KG', 'sum'),
                velikost_kartonu=('Box_Size', 'first')
            ).reset_index()

            top_100 = top_materials.sort_values(by='celkem_pohybu', ascending=False).head(100)
            
            top_100['velikost_kartonu'] = top_100['velikost_kartonu'].apply(lambda x: int(x) if x > 1 else t('val_loose'))

            top_100.rename(columns={
                'Material': t('col_mat'),
                'pocet_picku': t('col_lines'),
                'velikost_kartonu': t('col_box'),
                'celkem_pohybu': t('col_mov'),
                'celkove_mnozstvi': t('col_qty'),
                'celkova_natacena_vaha': t('col_wgt')
            }, inplace=True)

            # Export do Excelu (.xlsx) pomocí bufferu v paměti
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                top_100.to_excel(writer, index=False, sheet_name='TOP_100_Materials')
            
            # Zobrazení tlačítka pro stažení originálního .xlsx
            st.download_button(
                label=t('btn_download'),
                data=buffer.getvalue(),
                file_name="TOP_100_materialy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Zobrazení na obrazovce
            col_top1, col_top2 = st.columns([1.5, 1])

            with col_top1:
                st.dataframe(top_100.style.format({
                    t('col_wgt'): "{:.1f}",
                    t('col_mov'): "{:.0f}"
                }), use_container_width=True, hide_index=True)

            with col_top2:
                st.bar_chart(top_100.set_index(t('col_mat'))[t('col_mov')])

if __name__ == "__main__":
    main()
