import streamlit as st
import pandas as pd
import numpy as np
import io
import re

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
        'desc': "Tento analytický nástroj měří **skutečnou fyzickou náročnost vychystávání** ve skladu. Spojením provozních dat (Pick), kmenových dat (MARM) a **ručního ověření balení** eliminuje zkreslení počtu kusů a přesně modeluje fyzické pohyby (hmaty) u regálu.",
        'upload_title': "📁 Nahrání vstupních dat",
        'upload_help': "Nahrajte Pick report, MARM report a volitelně i soubor s ručním ověřením TOP materiálů.",
        'loading': "Načítám, identifikuji a propojuji soubory...",
        'err_pick': "Nepodařilo se najít Pick report (chybí sloupec 'Delivery'). Nahrajte prosím správný soubor.",
        'warn_marm': "Nebyl nahrán MARM report. Výpočet pohybů bude pouze orientační bez krabic a vah.",
        'info_clean': "💡 **Informace o čištění dat:** Z výpočtů bylo kompletně vyloučeno **{} zakázek**, u kterých došlo k odběru celé manipulační jednotky ('X').",
        'info_manual': "✅ **Ruční ověření:** Úspěšně načteno vlastní balení z nahrávaného souboru pro **{} materiálů**. Tato data mají přednost před MARMem!",
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
1. **Odběr po celých kartonech:** Systém primárně využívá ručně nahrané balení, jinak použije MARM. Dokáže počítat i s více krabicemi (např. nejprve spočítá krabice po 90 ks, pak po 6 ks = co balení, to 1 pohyb).
2. **Zbylé volné kusy (Těžké/Velké):** Zbylé kusy se zkontrolují podle váhy a rozměrů. Pokud kus váží více než limit **NEBO** je jeho nejdelší strana delší než limit, picker je musí brát po jednom = 1 kus je 1 pohyb.
3. **Zbylé volné kusy (Lehké/Malé):** U lehkých a drobných dílů systém předpokládá nabrání do hrsti (nastaveno v panelu) = např. 20 ks znamená jen cca 7 pohybů.
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
        'col_box': "Ks v balení (Hierarchie)",
        'val_loose': "Volné (Po kusech)",
        'btn_download': "📥 Stáhnout tabulku jako Excel (.xlsx)"
    },
    'en': {
        'switch_lang': "🇨🇿 Přepnout do češtiny",
        'title': "📈 Warehouse Workload & Picking Analysis",
        'desc': "This tool measures the **true physical demand of picking**. By merging Pick reports, MARM data, and **manual packaging overrides**, it accurately models how many physical hand movements the picker had to perform.",
        'upload_title': "📁 Upload Input Data",
        'upload_help': "Upload Pick report, MARM report, and optionally a Manual Override file for TOP materials.",
        'loading': "Loading, identifying, and merging files...",
        'err_pick': "Could not find Pick report (missing 'Delivery' column). Please upload the correct file.",
        'warn_marm': "MARM report not uploaded. Movement calculation will be approximate.",
        'info_clean': "💡 **Data Cleaning Info:** Excluded **{} orders** completely because they contained full handling unit picks ('X').",
        'info_manual': "✅ **Manual Override:** Successfully loaded custom packaging for **{} materials** from the uploaded file. These override MARM data!",
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
1. **Full Carton Picks:** The system primarily uses your manually verified packaging, otherwise falls back to MARM. It can handle multiple box sizes (e.g., 90 pcs boxes first, then 6 pcs boxes = 1 movement per box).
2. **Remaining Loose Pieces (Heavy/Large):** The remaining pieces are checked. If a piece exceeds the weight limit **OR** its longest side exceeds the dimension limit, they are picked one by one = 1 movement per piece.
3. **Remaining Loose Pieces (Light/Small):** For light and small parts, the system assumes grab-picking (e.g., 20 pcs equal approx 7 movements).
        """,
        'exp_detail_title': "Show details of filtered orders",
        'col_mat': "Material",
        'col_qty': "Total Pieces",
        'col_mov': "Hand Movements",
        'col_wgt': "Est. Weight (kg)",
        'col_max_dim': "Max Dim. (cm)",
        'col_cert': "Certificates",
        'no_orders': "No orders found matching the criteria.",
        'sec2_title': "🏆 TOP 100 Physically Most Demanding Materials",
        'col_lines': "Lines (Visits)",
        'col_box': "Pcs in Box (Hierarchy)",
        'val_loose': "Loose (Piece by piece)",
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
        df_manual = None

        with st.spinner(t('loading')):
            for file in uploaded_files:
                if file.name.lower().endswith('.csv'):
                    temp_df = pd.read_csv(file, dtype=str)
                else:
                    temp_df = pd.read_excel(file, dtype=str)
                
                # Inteligentní rozeznání souborů
                if 'Delivery' in temp_df.columns:
                    df_pick = temp_df
                elif 'Numerator' in temp_df.columns and 'Alternative Unit of Measure' in temp_df.columns:
                    df_marm = temp_df
                else:
                    # Pokud má aspoň 2 sloupce a není to ani MARM ani Pick, jde o soubor ručního ověření
                    if len(temp_df.columns) >= 2:
                        df_manual = temp_df

        if df_pick is None:
            st.error(t('err_pick'))
            return
            
        if df_marm is None:
            st.warning(t('warn_marm'))

        df_pick = df_pick.dropna(subset=['Delivery', 'Material']).copy()
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)

        # 1. VYŘAZENÍ ZAKÁZEK S 'X'
        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().str.upper()
        zakazky_s_x = df_pick[df_pick['Removal of total SU'] == 'X']['Delivery'].unique()
        df_pick = df_pick[~df_pick['Delivery'].isin(zakazky_s_x)].copy()
        
        st.info(t('info_clean').format(len(zakazky_s_x)))

        # 2. ZPRACOVÁNÍ RUČNÍHO OVĚŘENÍ MATERIÁLŮ (Priorita č.1)
        manual_boxes = {}
        if df_manual is not None and not df_manual.empty:
            col_mat = df_manual.columns[0]
            col_pkg = df_manual.columns[1]
            
            for idx, row in df_manual.iterrows():
                mat = str(row[col_mat]).strip()
                pkg = str(row[col_pkg]).strip()
                if pd.isna(mat) or mat == 'nan' or mat == 'None': continue
                
                # Chytrá Regex extrakce - najde všechna čísla před "ks" nebo za "K-"
                nums = re.findall(r'(\d+)\s*ks|\bK-(\d+)\b', pkg, flags=re.IGNORECASE)
                extracted = []
                for match in nums:
                    for group in match:
                        if group: extracted.append(int(group))
                
                # Seřadíme sestupně (od největší krabice po nejmenší pod-krabičky)
                extracted = sorted(extracted, reverse=True)
                
                # Pokud v textu nebylo číslo, ale je tam "po kusech", nastavíme nuceně krabici 1
                if not extracted and 'po kusech' in pkg.lower():
                    extracted = [1]
                    
                if extracted:
                    manual_boxes[mat] = extracted

            if manual_boxes:
                st.success(t('info_manual').format(len(manual_boxes)))

        # 3. ZPRACOVÁNÍ MARM DAT (Váhy, rozměry a záložní kartony)
        box_dict = {}
        weight_dict = {}
        dim_dict = {}

        if df_marm is not None:
            df_boxes = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK', 'KAR', 'KART', 'PAK', 'VPE', 'CAR', 'BLO'])]
            df_boxes['Numerator'] = pd.to_numeric(df_boxes['Numerator'], errors='coerce').fillna(0)
            box_sizes = df_boxes.groupby('Material')['Numerator'].max().to_dict()
            box_dict = {mat: int(size) for mat, size in box_sizes.items() if size > 1}

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

        # 4. NAPOJENÍ HIERARCHIE BALENÍ NA PICK REPORT
        def get_box_sizes(mat):
            if mat in manual_boxes:
                return manual_boxes[mat] # Použije hierarchii z ručního souboru (např. [90, 6])
            else:
                marm_b = box_dict.get(mat, 0)
                return [marm_b] if marm_b > 1 else [] # Záložní z MARM

        df_pick['Box_Sizes_List'] = df_pick['Material'].apply(get_box_sizes)
        df_pick['Piece_Weight_KG'] = df_pick['Material'].map(weight_dict).fillna(0)
        df_pick['Piece_Max_Dim_CM'] = df_pick['Material'].map(dim_dict).fillna(0)

        # POSTRANNÍ PANEL
        st.sidebar.header(t('sidebar_title'))
        limit_vahy = st.sidebar.number_input(
            t('weight_label'), min_value=0.1, max_value=20.0, value=2.0, step=0.5, help=t('weight_help'))
        limit_rozmeru = st.sidebar.number_input(
            t('dim_label'), min_value=1.0, max_value=200.0, value=15.0, step=1.0, help=t('dim_help'))
        kusy_na_hmat = st.sidebar.slider(
            t('hmat_label'), min_value=1, max_value=20, value=3, step=1, help=t('hmat_help'))
        st.sidebar.divider()
        unique_materials = sorted(df_pick['Material'].unique().tolist())
        excluded_materials = st.sidebar.multiselect(
            t('exclude_label'), options=unique_materials, default=[])
        
        if excluded_materials:
            df_pick = df_pick[~df_pick['Material'].isin(excluded_materials)]

        # 5. CHYTRÝ VÝPOČET POHYBŮ S HIERARCHIÍ BALENÍ
        def spocitej_pohyby(row):
            qty = row['Qty']
            if qty <= 0: return 0
            
            pohyby = 0
            zbytek = qty
            boxes = row['Box_Sizes_List']
            
            # Postupně odečítá kartony (od největší po nejmenší pod-krabice)
            for box_size in boxes:
                if box_size > 1 and zbytek >= box_size:
                    plne_kartony = zbytek // box_size
                    pohyby += plne_kartony
                    zbytek = zbytek % box_size
                
            # Na úplný zbytek volných kusů kontroluje váhu a rozměry
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

        st.divider()
        st.subheader(t('sec1_title'))
        
        if len(filtered_orders) > 0:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric(t('m_orders'), f"{len(filtered_orders):,}".replace(',', ' '))
            col2.metric(t('m_qty'), f"{filtered_orders['total_qty'].mean():.1f}")
            col3.metric(t('m_pos'), f"{filtered_orders['num_positions'].mean():.2f}")
            col4.metric(t('m_mov'), f"{filtered_orders['celkem_pohybu'].mean():.1f}")

            with st.expander(t('exp_calc_title')):
                st.markdown(t('exp_calc_text'))

            with st.expander(t('exp_detail_title')):
                display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'vaha_zakazky', 'max_rozmer', 'certs']].copy()
                display_df.columns = [t('col_mat'), t('col_qty'), t('col_mov'), t('col_wgt'), t('col_max_dim'), t('col_cert')]
                st.dataframe(display_df, use_container_width=True)
        else:
            st.warning(t('no_orders'))

        # SEKCE 2: TOP 100 NEJNÁROČNĚJŠÍCH MATERIÁLŮ
        st.divider()
        st.subheader(t('sec2_title'))
        
        if not df_pick.empty:
            # Pomocná funkce pro vypsání hierarchie balení do tabulky (např. "90ks + 6ks")
            def format_box_sizes(boxes):
                if not boxes or boxes == [1]:
                    return t('val_loose')
                return " + ".join([f"{b}ks" for b in boxes])

            top_materials = df_pick.groupby('Material').agg(
                pocet_picku=('Material', 'count'),
                celkove_mnozstvi=('Qty', 'sum'),
                celkem_pohybu=('Pohyby_Rukou', 'sum'),
                celkova_natacena_vaha=('Celkova_Vaha_KG', 'sum'),
                Box_Sizes_List=('Box_Sizes_List', 'first')
            ).reset_index()

            top_100 = top_materials.sort_values(by='celkem_pohybu', ascending=False).head(100)
            top_100['velikost_kartonu'] = top_100['Box_Sizes_List'].apply(format_box_sizes)
            top_100 = top_100.drop(columns=['Box_Sizes_List'])

            top_100.columns = [t('col_mat'), t('col_lines'), t('col_qty'), t('col_mov'), t('col_wgt'), t('col_box')]
            
            # Přehození sloupců pro logičtější pořadí
            top_100 = top_100[[t('col_mat'), t('col_lines'), t('col_box'), t('col_qty'), t('col_wgt'), t('col_mov')]]

            # Export do Excelu (.xlsx) pomocí bufferu v paměti
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                top_100.to_excel(writer, index=False, sheet_name='TOP_100_Materials')
            
            st.download_button(
                label=t('btn_download'),
                data=buffer.getvalue(),
                file_name="TOP_100_materialy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            col_top1, col_top2 = st.columns([1.5, 1])
            with col_top1:
                st.dataframe(top_100.style.format({t('col_wgt'): "{:.1f}", t('col_mov'): "{:.0f}"}), use_container_width=True, hide_index=True)
            with col_top2:
                st.bar_chart(top_100.set_index(t('col_mat'))[t('col_mov')])

if __name__ == "__main__":
    main()
