import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl.chart import BarChart, Reference

# ==========================================
# 1. NASTAVENÍ A LOKALIZACE
# ==========================================
st.set_page_config(page_title="Analýza Skladové Zátěže", page_icon="📈", layout="wide")

if 'lang' not in st.session_state:
    st.session_state.lang = 'cs'

TEXTS = {
    'cs': {
        'switch_lang': "🇬🇧 Switch to English",
        'title': "📈 Analýza Skladové Zátěže a Ergonomie Pickování",
        'desc': "Nástroj pro měření **skutečné fyzické náročnosti** vychystávání. Tato aplikace modeluje **skutečný počet pohybů rukou (hmatů)** a **zvednutou hmotnost** u regálu.",
        'upload_title': "📁 Nahrání vstupních dat",
        'upload_help': "Nahrajte Pick report, MARM report, TO details (Queue) a volitelně i ruční ověření balení.",
        'loading': "Zpracovávám logiku hierarchie balení a propojuji s Queue...",
        'err_pick': "Nepodařilo se najít Pick report (chybí sloupec 'Delivery' nebo 'Act.qty (dest)').",
        'info_clean': "💡 Z výpočtů jemné motoriky bylo vyloučeno **{} zakázek**, kde byla snímána celá manipulační jednotka / paleta ('X').",
        'info_manual': "✅ Načteno ruční ověření balení pro **{} materiálů**. (Priorita před MARM)",
        'sidebar_title': "⚙️ Fyzické limity pickera",
        'weight_label': "Hranice pro nošení po 1 ks (kg)",
        'dim_label': "Hranice rozměru pro 1 ks (cm)",
        'hmat_label': "Max ks lehkých dílů do hrsti",
        'exclude_label': "Vyloučit materiály z výpočtů:",
        'sec_methodology': "📖 Pro management: Jak a proč se počítají pohyby?",
        'methodology_text': """
### ⚙️ Krok za krokem: Jak funguje algoritmus?
1. **Zjistí balení (Kartony):** Nejdříve se podívá do ručních ověření, poté do MARMu. Zjistí, zda se materiál nachází v krabici (např. 50 ks). Pokud picker vychystává 120 ks, započítá odběr **2 celých krabic = 2 pohyby**. Zbyde 20 volných kusů.
2. **Vyhodnotí váhu a rozměr zbytku:** U zbylých 20 ks zkontroluje limity (např. >2 kg nebo >15 cm). Pokud kus limit překračuje, musí se brát po jednom kusu = **20 pohybů**.
3. **Drobné díly do hrsti:** Pokud jsou kusy naopak lehké a malé, předpokládáme nabrání do hrsti (např. 3 ks na hmat) = **7 pohybů**.
        """,
        'sec_ratio': "🎯 Zdroj výpočtů (Spolehlivost dat)",
        'ratio_desc': "Tento přehled ukazuje, jak kvalitní data jsme měli pro výpočet. Pohyby se dělí na přesně identifikované krabice, na přirozené zbytky/ověřené volné kusy a na odhady.",
        'ratio_master': "Přesně (Krabice / Pytlíky)",
        'ratio_loose_ok': "Přesně (Ověřené volné / Zbytky)",
        'ratio_loose_miss': "Odhad (Chybí data o balení)",
        'sec_queue_title': "📊 Průměrná náročnost dle typu pickování (Queue)",
        'q_col_queue': "Typ Pickování (Queue)",
        'q_col_orders': "Počet zakázek",
        'q_col_pcs': "Prům. kusů",
        'q_col_moves': "Prům. celkem pohybů",
        'q_col_box': "Prům. krabice/pytlíky",
        'q_col_ok': "Prům. ověřené volné",
        'q_col_miss': "Prům. chybí balení",
        'sec1_title': "🎯 Analýza paletových zakázek (Obsahují pouze 1 materiál)",
        'm_orders': "Počet zakázek",
        'm_qty': "Prům. kusů / zakázku",
        'm_pos': "Prům. pozic / zakázku",
        'm_mov': "Prům. fyz. pohybů",
        'exp_detail_title': "Zobrazit tabulku zakázek (1 materiál)",
        'col_mat': "Materiál",
        'col_qty': "Kusů celkem",
        'col_mov': "Celkem pohybů",
        'col_mov_box': "Pohyby (Krabice/Pytlíky)",
        'col_mov_loose_ok': "Pohyby (Ověřené volné)",
        'col_mov_loose_miss': "Pohyby (Chybí balení)",
        'col_wgt': "Hmotnost (kg)",
        'col_max_dim': "Rozměr (cm)",
        'col_cert': "Certifikát",
        'sec2_title': "🏆 TOP 100 fyzicky nejnáročnějších materiálů (dle hmatů)",
        'col_lines': "Řádky (Návštěvy)",
        'col_box': "Hierarchie balení",
        'val_loose': "Volné kusy",
        'btn_download': "📥 Stáhnout kompletní report (Excel Workbook)",
        'no_orders': "Nenalezeny žádné zakázky pro zobrazení.",
        'sec3_title': "🔍 Vyhledávač Master Dat materiálu",
        'search_label': "Vyhledejte a vyberte materiál pro zobrazení detailů ze SAPu a ručního ověření:"
    },
    'en': {
        'switch_lang': "🇨🇿 Přepnout do češtiny",
        'title': "📈 Warehouse Workload & Ergonomics Analysis",
        'desc': "A tool to measure the **true physical demand** of picking. This app merges SAP (MARM) and manual verifications to model **actual hand movements** and **lifted weight**.",
        'upload_title': "📁 Upload Input Data",
        'upload_help': "Upload Pick report, MARM report, TO details (Queue), and optional Manual Override file.",
        'loading': "Processing packaging hierarchy and Queue mapping...",
        'err_pick': "Pick report missing (no 'Delivery' or 'Act.qty (dest)' column).",
        'info_clean': "💡 Excluded **{} orders** consisting of full handling units ('X').",
        'info_manual': "✅ Loaded manual packaging for **{} materials**. (Overrides MARM)",
        'sidebar_title': "⚙️ Picker's Physical Limits",
        'weight_label': "Weight limit for 1-by-1 pick (kg)",
        'dim_label': "Dimension limit for 1-by-1 (cm)",
        'hmat_label': "Max pieces per grab (light parts)",
        'exclude_label': "Exclude materials:",
        'sec_methodology': "📖 For Management: Why and how do we calculate movements?",
        'methodology_text': """
### ⚙️ Step by Step: How does the algorithm work?
1. **Identify Cartons:** Checks manual data, then MARM. If a box holds 50 pcs and order is 120 pcs -> **2 full boxes = 2 movements**. 20 pcs remain.
2. **Evaluate Heavy/Large:** If the remaining 20 pcs exceed the weight/dimension limit, they are picked individually = **20 movements**.
3. **Small Handfuls:** If light/small, we assume grabbing (e.g. 3 pcs/grab) = **7 movements**.
        """,
        'sec_ratio': "🎯 Calculation Source (Data Reliability)",
        'ratio_desc': "Shows how movements were calculated: exact full boxes, exact verified loose, and missing box data.",
        'ratio_master': "Exact (Boxes / Bags)",
        'ratio_loose_ok': "Exact (Verified Loose / Remainders)",
        'ratio_loose_miss': "Estimated (Missing Box Data)",
        'sec_queue_title': "📊 Average Workload by Picking Type (Queue)",
        'q_col_queue': "Picking Type (Queue)",
        'q_col_orders': "Orders Count",
        'q_col_pcs': "Avg Pieces",
        'q_col_moves': "Avg Total Moves",
        'q_col_box': "Avg Boxes/Bags",
        'q_col_ok': "Avg Verified Loose",
        'q_col_miss': "Avg Missing Box",
        'sec1_title': "🎯 Single-Material Pallet Orders",
        'm_orders': "Orders",
        'm_qty': "Avg Pcs / Order",
        'm_pos': "Avg Bins / Order",
        'm_mov': "Avg Physical Moves",
        'exp_detail_title': "Show Orders Table (1 Material)",
        'col_mat': "Material",
        'col_qty': "Total Pieces",
        'col_mov': "Total Movements",
        'col_mov_box': "Moves (Boxes / Bags)",
        'col_mov_loose_ok': "Moves (Verified Loose)",
        'col_mov_loose_miss': "Moves (Missing Box)",
        'col_wgt': "Weight (kg)",
        'col_max_dim': "Max Dim (cm)",
        'col_cert': "Certificate",
        'sec2_title': "🏆 TOP 100 Most Demanding Materials (by movements)",
        'col_lines': "Lines (Visits)",
        'col_box': "Packaging Hierarchy",
        'val_loose': "Loose",
        'btn_download': "📥 Download Comprehensive Report (Excel Workbook)",
        'no_orders': "No orders found.",
        'sec3_title': "🔍 Master Data Material Viewer",
        'search_label': "Search and select a material to show its SAP and manual override details:"
    }
}

def t(key): return TEXTS[st.session_state.lang][key]

# SKRYTÝ PŘEKLADAČ MATERIÁLŮ (Match Key) - řeší problém excelových nul bez toho, aby změnil vzhled původního čísla
def get_match_key(val):
    v = str(val).strip().upper()
    if '.' in v and v.replace('.', '').isdigit():
        return v.rstrip('0').rstrip('.')
    return v

def main():
    col_spacer, col_lang = st.columns([8, 1])
    with col_lang:
        if st.button(t('switch_lang')):
            st.session_state.lang = 'en' if st.session_state.lang == 'cs' else 'cs'
            st.rerun()

    st.title(t('title'))
    st.markdown(t('desc'))
    
    with st.expander(t('sec_methodology'), expanded=False):
        st.markdown(t('methodology_text'))
    st.divider()

    with st.container():
        st.subheader(t('upload_title'))
        uploaded_files = st.file_uploader(t('upload_help'), type=['csv', 'xlsx'], accept_multiple_files=True)

    if uploaded_files:
        df_pick, df_marm, df_manual, df_queue = None, None, None, None

        with st.spinner(t('loading')):
            for file in uploaded_files:
                if file.name.lower().endswith('.csv'): 
                    temp_df = pd.read_csv(file, dtype=str)
                else: 
                    temp_df = pd.read_excel(file, dtype=str)
                
                # Detekce souborů
                if 'Delivery' in temp_df.columns and 'Act.qty (dest)' in temp_df.columns: 
                    df_pick = temp_df
                elif 'Numerator' in temp_df.columns and 'Alternative Unit of Measure' in temp_df.columns: 
                    df_marm = temp_df
                elif 'Queue' in temp_df.columns and ('Transfer Order Number' in temp_df.columns or 'SD Document' in temp_df.columns):
                    df_queue = temp_df
                else:
                    if len(temp_df.columns) >= 2: 
                        df_manual = temp_df

        if df_pick is None:
            st.error(t('err_pick'))
            return

        # 1. PŘÍPRAVA PICK REPORTU (ZACHOVÁNÍ PŮVODNÍCH ČÍSEL)
        df_pick['Material'] = df_pick['Material'].astype(str).str.strip()
        df_pick['Match_Key'] = df_pick['Material'].apply(get_match_key) # Skrytý klíč pro propojování
        
        df_pick = df_pick.dropna(subset=['Delivery', 'Material']).copy()
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)

        # Propojení s TOs details a rozdělení dle Queue
        if df_queue is not None:
            if 'SD Document' in df_queue.columns:
                q_map = df_queue.dropna(subset=['SD Document', 'Queue']).drop_duplicates('SD Document').set_index('SD Document')['Queue'].to_dict()
                df_pick['Queue'] = df_pick['Delivery'].map(q_map)

        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().str.upper()
        zakazky_s_x = df_pick[df_pick['Removal of total SU'] == 'X']['Delivery'].unique()
        df_pick = df_pick[~df_pick['Delivery'].isin(zakazky_s_x)].copy()
        st.info(t('info_clean').format(len(zakazky_s_x)))

        # 2. NAČTENÍ RUČNÍCH BALENÍ
        manual_boxes = {}
        if df_manual is not None and not df_manual.empty:
            c_mat, c_pkg = df_manual.columns[0], df_manual.columns[1]
            for _, row in df_manual.iterrows():
                mat_raw = str(row[c_mat]).strip()
                if pd.isna(mat_raw) or mat_raw.upper() in ['NAN', 'NONE', '']: continue
                
                mat_key = get_match_key(mat_raw) # Přeložíme přes skrytý klíč!
                pkg = str(row[c_pkg]).strip()
                
                nums = re.findall(r'(\d+)\s*ks|\bK-(\d+)\b|(?:pytl[íi]k|pytel|role|balen[íi]|krabice)[^\d]*(\d+)', pkg, flags=re.IGNORECASE)
                ext = sorted(list(set([int(g) for m in nums for g in m if g])), reverse=True)
                
                if not ext and 'po kusech' in pkg.lower(): 
                    ext = [1]
                    
                if ext: 
                    manual_boxes[mat_key] = ext
                    
            if manual_boxes: st.success(t('info_manual').format(len(manual_boxes)))

        # 3. NAČTENÍ MARM
        box_dict, weight_dict, dim_dict = {}, {}, {}
        if df_marm is not None:
            df_marm['Match_Key'] = df_marm['Material'].apply(get_match_key)
            
            df_boxes = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK', 'KAR', 'KART', 'PAK', 'VPE', 'CAR', 'BLO'])]
            df_boxes['Numerator'] = pd.to_numeric(df_boxes['Numerator'], errors='coerce').fillna(0)
            
            def get_sorted_boxes(group):
                return sorted([int(x) for x in group if x > 1], reverse=True)
            box_dict = df_boxes.groupby('Match_Key')['Numerator'].apply(get_sorted_boxes).to_dict()

            df_st = df_marm[df_marm['Alternative Unit of Measure'].isin(['ST', 'PCE', 'KS'])].copy()
            df_st['Gross Weight'] = pd.to_numeric(df_st['Gross Weight'], errors='coerce').fillna(0)
            df_st['Weight_KG'] = df_st.apply(lambda r: r['Gross Weight']/1000.0 if str(r['Unit of Weight']).upper()=='G' else r['Gross Weight'], axis=1)
            weight_dict = df_st.groupby('Match_Key')['Weight_KG'].first().to_dict()

            def to_cm(val, unit):
                try:
                    v, u = float(val), str(unit).upper().strip()
                    if u == 'MM': return v / 10.0
                    if u == 'M': return v * 100.0
                    return v 
                except: return 0.0

            df_st['L'] = df_st.apply(lambda r: to_cm(r['Length'], r['Unit of Dimension']), axis=1)
            df_st['W'] = df_st.apply(lambda r: to_cm(r['Width'], r['Unit of Dimension']), axis=1)
            df_st['H'] = df_st.apply(lambda r: to_cm(r['Height'], r['Unit of Dimension']), axis=1)
            df_st['Max_Dim_CM'] = df_st[['L', 'W', 'H']].max(axis=1) 
            dim_dict = df_st.groupby('Match_Key')['Max_Dim_CM'].first().to_dict()

        # Připojení dat přes MATCH KEY! (takže originální Material column zůstane nenarušen)
        df_pick['Box_Sizes_List'] = df_pick['Match_Key'].apply(lambda m: manual_boxes.get(m, box_dict.get(m, [])))
        df_pick['Piece_Weight_KG'] = df_pick['Match_Key'].map(weight_dict).fillna(0)
        df_pick['Piece_Max_Dim_CM'] = df_pick['Match_Key'].map(dim_dict).fillna(0)

        # Postranní panel
        st.sidebar.header(t('sidebar_title'))
        limit_vahy = st.sidebar.number_input(t('weight_label'), min_value=0.1, max_value=20.0, value=2.0, step=0.5)
        limit_rozmeru = st.sidebar.number_input(t('dim_label'), min_value=1.0, max_value=200.0, value=15.0, step=1.0)
        kusy_na_hmat = st.sidebar.slider(t('hmat_label'), min_value=1, max_value=20, value=3, step=1)
        
        st.sidebar.divider()
        unique_materials = sorted(df_pick['Material'].unique().tolist()) # Pro dropdown bereme ty hezké z SAPu!
        excluded_materials = st.sidebar.multiselect(t('exclude_label'), options=unique_materials, default=[])
        if excluded_materials: df_pick = df_pick[~df_pick['Material'].isin(excluded_materials)]

        # ==========================================
        # 4. VÝPOČET POHYBŮ S ROZPADEM
        # ==========================================
        def spocitej_pohyby_detail(row):
            qty = row['Qty']
            if qty <= 0: return 0, 0, 0, 0
            
            pohyby_box, pohyby_loose_ok, pohyby_loose_miss = 0, 0, 0
            zbytek = qty
            boxes = row['Box_Sizes_List']
            
            for box_size in boxes:
                if box_size > 1 and zbytek >= box_size:
                    pohyby_box += zbytek // box_size
                    zbytek = zbytek % box_size
                
            if zbytek > 0:
                if row['Piece_Weight_KG'] >= limit_vahy or row['Piece_Max_Dim_CM'] >= limit_rozmeru:
                    pohyby = zbytek
                else:
                    pohyby = np.ceil(zbytek / kusy_na_hmat)
                
                if len(boxes) == 0: pohyby_loose_miss += pohyby
                else: pohyby_loose_ok += pohyby
                    
            return pohyby_box + pohyby_loose_ok + pohyby_loose_miss, pohyby_box, pohyby_loose_ok, pohyby_loose_miss

        df_pick[['Pohyby_Rukou', 'Pohyby_Box', 'Pohyby_Loose_OK', 'Pohyby_Loose_Miss']] = df_pick.apply(spocitej_pohyby_detail, axis=1, result_type='expand')
        df_pick['Celkova_Vaha_KG'] = df_pick['Qty'] * df_pick['Piece_Weight_KG']

        # ==========================================
        # ZOBRAZENÍ POMĚRU A QUEUE
        # ==========================================
        total_pohyby = df_pick['Pohyby_Rukou'].sum()
        total_box = df_pick['Pohyby_Box'].sum()
        total_loose_ok = df_pick['Pohyby_Loose_OK'].sum()
        total_loose_miss = df_pick['Pohyby_Loose_Miss'].sum()

        if total_pohyby > 0:
            st.divider()
            st.subheader(t('sec_ratio'))
            st.markdown(t('ratio_desc'))
            
            pct_box = (total_box / total_pohyby) * 100
            pct_ok = (total_loose_ok / total_pohyby) * 100
            pct_miss = (total_loose_miss / total_pohyby) * 100
            
            col_r1, col_r2, col_r3 = st.columns(3)
            col_r1.metric(t('ratio_master'), f"{pct_box:.1f} %", f"{total_box:,.0f} pohybů")
            col_r2.metric(t('ratio_loose_ok'), f"{pct_ok:.1f} %", f"{total_loose_ok:,.0f} pohybů")
            col_r3.metric(t('ratio_loose_miss'), f"{pct_miss:.1f} %", f"{total_loose_miss:,.0f} pohybů", delta_color="inverse")

        # SEKCE: QUEUE (Pokud je TO's_details nahrán)
        queue_summary = None
        if 'Queue' in df_pick.columns and df_pick['Queue'].notna().any():
            st.divider()
            st.subheader(t('sec_queue_title'))
            
            queue_agg = df_pick.groupby(['Delivery', 'Queue']).agg(
                celkem_pohybu=('Pohyby_Rukou', 'sum'),
                pohyby_box=('Pohyby_Box', 'sum'),
                pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'),
                pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
                total_qty=('Qty', 'sum')
            ).reset_index()
            
            queue_summary = queue_agg.groupby('Queue').agg(
                pocet_zakazek=('Delivery', 'nunique'),
                prum_kusu=('total_qty', 'mean'),
                prum_pohybu=('celkem_pohybu', 'mean'),
                prum_box=('pohyby_box', 'mean'),
                prum_ok=('pohyby_loose_ok', 'mean'),
                prum_miss=('pohyby_loose_miss', 'mean')
            ).reset_index().sort_values('prum_pohybu', ascending=False)
            
            queue_summary.columns = [
                t('q_col_queue'), t('q_col_orders'), t('q_col_pcs'), 
                t('q_col_moves'), t('q_col_box'), t('q_col_ok'), t('q_col_miss')
            ]
            
            st.dataframe(queue_summary.style.format({
                t('q_col_pcs'): "{:.1f}", 
                t('q_col_moves'): "{:.1f}",
                t('q_col_box'): "{:.1f}",
                t('q_col_ok'): "{:.1f}",
                t('q_col_miss'): "{:.1f}"
            }), use_container_width=True, hide_index=True)

        # ==========================================
        # ZAKÁZKY A TOP 100
        # ==========================================
        all_materials_agg = df_pick.groupby('Material').agg(
            pocet_picku=('Material', 'count'),
            celkove_mnozstvi=('Qty', 'sum'),
            celkem_pohybu=('Pohyby_Rukou', 'sum'),
            pohyby_box=('Pohyby_Box', 'sum'),
            pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'),
            pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
            celkova_natacena_vaha=('Celkova_Vaha_KG', 'sum'),
            Box_Sizes_List=('Box_Sizes_List', 'first')
        ).reset_index()

        all_materials_agg['velikost_kartonu'] = all_materials_agg['Box_Sizes_List'].apply(
            lambda b: " + ".join([f"{x}ks" for x in b]) if b and b != [1] else t('val_loose'))

        all_materials_agg.rename(columns={
            'Material': t('col_mat'),
            'pocet_picku': t('col_lines'),
            'velikost_kartonu': t('col_box'),
            'celkem_pohybu': t('col_mov'),
            'pohyby_box': t('col_mov_box'),
            'pohyby_loose_ok': t('col_mov_loose_ok'),
            'pohyby_loose_miss': t('col_mov_loose_miss'),
            'celkove_mnozstvi': t('col_qty'),
            'celkova_natacena_vaha': t('col_wgt')
        }, inplace=True)

        top_100 = all_materials_agg.sort_values(by=t('col_mov'), ascending=False).head(100)
        top_100 = top_100[[t('col_mat'), t('col_lines'), t('col_box'), t('col_qty'), t('col_wgt'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_mov')]]

        def is_valid_cert(certs):
            valid_certs = [str(c).strip() for c in certs if pd.notna(c) and str(c).strip() not in ['nan', '']]
            if len(valid_certs) == 0: return False
            for c in valid_certs:
                if c.startswith('460'): return False
            return True

        grouped_orders = df_pick.groupby('Delivery').agg(
            num_materials=('Material', 'nunique'), material=('Material', 'first'),
            certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
            total_qty=('Qty', 'sum'), num_positions=('Source Storage Bin', 'nunique'),
            celkem_pohybu=('Pohyby_Rukou', 'sum'), pohyby_box=('Pohyby_Box', 'sum'), 
            pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'), pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
            vaha_zakazky=('Celkova_Vaha_KG', 'sum'), max_rozmer=('Piece_Max_Dim_CM', 'first')
        )
        filtered_orders = grouped_orders[(grouped_orders['num_materials'] == 1) & (grouped_orders['certs'].apply(is_valid_cert))]

        st.divider()
        st.subheader(t('sec1_title'))
        
        if len(filtered_orders) > 0:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(t('m_orders'), f"{len(filtered_orders):,}".replace(',', ' '))
            c2.metric(t('m_qty'), f"{filtered_orders['total_qty'].mean():.1f}")
            c3.metric(t('m_pos'), f"{filtered_orders['num_positions'].mean():.2f}")
            c4.metric(t('m_mov'), f"{filtered_orders['celkem_pohybu'].mean():.1f}")

            with st.expander(t('exp_detail_title')):
                display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'pohyby_box', 'pohyby_loose_ok', 'pohyby_loose_miss', 'vaha_zakazky', 'max_rozmer', 'certs']].copy()
                display_df.columns = [t('col_mat'), t('col_qty'), t('col_mov'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_wgt'), t('col_max_dim'), t('col_cert')]
                st.dataframe(display_df, use_container_width=True)
        else:
            st.warning(t('no_orders'))

        # ==========================================
        # EXPORT DO EXCELU 
        # ==========================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            metodika_df = pd.DataFrame({
                "Téma": ["O reportu", "Nastavení (Hranice váhy)", "Nastavení (Hranice rozměru)", "Nastavení (Max do hrsti)"],
                "Popis": ["Report odstraňuje iluzi 'kusů' a odhaduje pohyby rukou.", f"{limit_vahy} kg", f"{limit_rozmeru} cm", f"{kusy_na_hmat} ks"]
            })
            metodika_df.to_excel(writer, index=False, sheet_name='Info_a_Metodika')
            
            # Export nové Queue sekce do vlastního listu
            if queue_summary is not None:
                queue_summary.to_excel(writer, index=False, sheet_name='Analyza_Queue')
                
            zakazky_export = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'pohyby_box', 'pohyby_loose_ok', 'pohyby_loose_miss', 'vaha_zakazky', 'max_rozmer']].copy()
            zakazky_export.columns = [t('col_mat'), t('col_qty'), t('col_mov'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_wgt'), t('col_max_dim')]
            zakazky_export.to_excel(writer, index=True, sheet_name='Souhrn_Zakazek')

            top_100.to_excel(writer, index=False, sheet_name='TOP_100_Materialy')
            
            workbook = writer.book
            worksheet = writer.sheets['TOP_100_Materialy']
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = "Zátěž materiálů dle fyzických pohybů"
            chart.y_axis.title = t('col_mov')
            chart.x_axis.title = t('col_mat')
            chart.width = 25
            chart.height = 12
            
            col_mat_idx = list(top_100.columns).index(t('col_mat')) + 1
            col_mov_idx = list(top_100.columns).index(t('col_mov')) + 1
            data = Reference(worksheet, min_col=col_mov_idx, min_row=1, max_row=len(top_100)+1)
            cats = Reference(worksheet, min_col=col_mat_idx, min_row=2, max_row=len(top_100)+1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.legend = None 
            worksheet.add_chart(chart, "K2")

            all_materials_export = all_materials_agg.drop(columns=['Box_Sizes_List'])
            all_materials_export.to_excel(writer, index=False, sheet_name='Vsechna_Data_Materialu')

        st.divider()
        st.subheader(t('sec2_title'))
        
        st.download_button(
            label=t('btn_download'),
            data=buffer.getvalue(),
            file_name="Analýza_Ergonomie_skladu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        col_top1, col_top2 = st.columns([1.5, 1])
        with col_top1:
            st.dataframe(top_100.style.format({t('col_wgt'): "{:.1f}", t('col_mov'): "{:.0f}", t('col_mov_box'): "{:.0f}", t('col_mov_loose_ok'): "{:.0f}", t('col_mov_loose_miss'): "{:.0f}"}), use_container_width=True, hide_index=True)
        with col_top2:
            st.bar_chart(top_100.set_index(t('col_mat'))[t('col_mov')])

        # ==========================================
        # PROHLÍŽEČ MASTER DAT
        # ==========================================
        st.divider()
        st.subheader(t('sec3_title'))
        
        mat_search = st.selectbox(t('search_label'), options=[""] + unique_materials)
        
        if mat_search:
            st.markdown(f"#### Detail pro materiál: **`{mat_search}`**")
            search_key = get_match_key(mat_search)
            
            if search_key in manual_boxes:
                if manual_boxes[search_key] == [1]:
                    st.success("✅ **Ruční ověření:** Nastaveno natvrdo jako **Volné kusy (1 ks)**.")
                else:
                    st.success(f"✅ **Ruční ověření nalezeno:** Nastaveny krabice/pytlíky po **{manual_boxes[search_key]} ks**.")
            else:
                st.info("ℹ️ Tento materiál nemá zadané žádné ruční ověření.")
                
            c_info1, c_info2 = st.columns(2)
            c_info1.metric("Váha 1 kusu (z MARM)", f"{weight_dict.get(search_key, 0):.3f} kg")
            c_info2.metric("Nejdelší rozměr (z MARM)", f"{dim_dict.get(search_key, 0):.1f} cm")
            
            if df_marm is not None:
                st.write("**Surová data z MARM reportu (Varianty balení):**")
                marm_detail = df_marm[df_marm['Match_Key'] == search_key]
                if not marm_detail.empty:
                    cols_to_show = ['Alternative Unit of Measure', 'Numerator', 'Denominator', 'Gross Weight', 'Unit of Weight', 'Length', 'Width', 'Height', 'Unit of Dimension']
                    available_cols = [c for c in cols_to_show if c in marm_detail.columns]
                    st.dataframe(marm_detail[available_cols], hide_index=True, use_container_width=True)
                else:
                    st.warning("⚠️ Pro tento materiál nebyla nalezena žádná data v MARM reportu.")

if __name__ == "__main__":
    main()
