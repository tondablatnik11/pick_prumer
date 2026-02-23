import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import time
from openpyxl.chart import BarChart, Reference

# ==========================================
# 1. NASTAVENÍ STRÁNKY A CSS VZHLEDU
# ==========================================
st.set_page_config(page_title="Skladová Analytika & Ergonomie", page_icon="📦", layout="wide")

st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        padding: 5% 5% 5% 10%;
        border-radius: 8px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    </style>
""", unsafe_allow_html=True)

if 'lang' not in st.session_state:
    st.session_state.lang = 'cs'

# ==========================================
# 2. LOKALIZACE TEXTŮ
# ==========================================
TEXTS = {
    'cs': {
        'switch_lang': "🇬🇧 Switch to English",
        'title': "📦 Skladová Analytika a Ergonomie",
        'desc': "Profesionální nástroj pro modelování skutečné fyzické zátěže pickování a validaci fakturačních procesů.",
        'upload_title': "📁 Nahrání vstupních dat (Klikněte pro sbalení/rozbalení)",
        'upload_help': "Nahrajte Pick report, MARM report, TO details (Queue) a volitelně i ruční ověření balení.",
        'info_users': "💡 Vyloučeno **{} systémových řádků** (UIDJ5089, UIH25501).",
        'info_clean': "💡 Detekováno **{} zakázek** s odběrem celé palety ('X'). (Započítán 1 pohyb).",
        'info_manual': "✅ Načteno ruční ověření pro **{} unikátních materiálů**.",
        'sidebar_title': "⚙️ Konfigurace algoritmů",
        'weight_label': "Hranice pro nošení po 1 ks (kg)",
        'dim_label': "Hranice rozměru pro 1 ks (cm)",
        'hmat_label': "Max ks lehkých dílů do hrsti",
        'exclude_label': "Vyloučit materiály z výpočtů:",
        'sec_ratio': "🎯 Spolehlivost dat a zdroj výpočtů",
        'ratio_desc': "Z jakých podkladů aplikace vycházela při výpočtu fyzických pohybů:",
        'ratio_master': "Přesně (Krabice / Palety)",
        'ratio_loose_ok': "Přesně (Ověřené volné)",
        'ratio_loose_miss': "Odhady (Chybí balení)",
        'exp_missing_data': "🔍 Zobrazit materiály s chybějícími daty o balení (Žebříček 'odhadů')",
        'sec_queue_title': "📊 Průměrná náročnost dle typu pickování (Queue)",
        'q_col_queue': "Queue",
        'q_col_to': "Počet TO",
        'q_col_orders': "Zakázky",
        'q_col_loc': "Prům. lokací",
        'q_col_pcs': "Prům. kusů",
        'q_col_moves': "Prům. pohybů na TO",
        'q_col_box': "Prům. krabice",
        'q_pct_box': "% Krabice",
        'q_col_ok': "Prům. volné",
        'q_pct_ok': "% Volné",
        'q_col_miss': "Prům. chybí",
        'q_pct_miss': "% Chybí",
        'sec_queue_top_title': "🏆 TOP 100 materiálů podle Queue",
        'q_select': "Zobrazit TOP 100 pro Queue:",
        'sec1_title': "🎯 Analýza paletových zakázek (1 materiál)",
        'm_orders': "Počet zakázek",
        'm_qty': "Prům. kusů / zakázku",
        'm_pos': "Prům. pozic / zakázku",
        'm_mov': "Prům. fyz. pohybů",
        'exp_detail_title': "Zobrazit tabulku zakázek (1 materiál)",
        'col_mat': "Materiál",
        'col_qty': "Kusů celkem",
        'col_mov': "Celkem pohybů",
        'col_mov_box': "Pohyby (Krabice)",
        'col_mov_loose_ok': "Pohyby (Volné)",
        'col_mov_loose_miss': "Pohyby (Chybí balení)",
        'col_wgt': "Hmotnost (kg)",
        'col_max_dim': "Rozměr (cm)",
        'sec1_top_title': "🏆 TOP 100 materiálů pro tyto paletové zakázky",
        'btn_download': "📥 Stáhnout kompletní report (Excel)",
        'no_orders': "Nenalezeny žádné zakázky pro zobrazení.",
        'sec3_title': "🔍 Prohlížeč Master Dat",
        'search_label': "Zkontrolujte si konkrétní materiál:",
        'tab_dashboard': "📊 Dashboard & Queue",
        'tab_pallets': "📦 Paletové zakázky",
        'tab_top': "🏆 TOP Materiály & Datová kvalita",
        'tab_audit': "🔍 Nástroje & Audit",
        'col_lines': "Řádky",
        'col_box': "Balení",
        'val_loose': "Volné kusy"
    },
    'en': {
        'switch_lang': "🇨🇿 Přepnout do češtiny",
        'title': "📦 Warehouse Analytics & Ergonomics",
        'desc': "Professional tool for modeling true physical picking workload and validating billing processes.",
        'upload_title': "📁 Upload Input Data (Click to expand/collapse)",
        'upload_help': "Upload Pick report, MARM report, TO details (Queue), and optional Manual Override.",
        'info_users': "💡 Excluded **{} system lines** (UIDJ5089, UIH25501).",
        'info_clean': "💡 Detected **{} full SU ('X') orders**. (Calculated as 1 move).",
        'info_manual': "✅ Loaded manual packaging for **{} unique materials**.",
        'sidebar_title': "⚙️ Algorithm Configuration",
        'weight_label': "Weight limit for 1-by-1 pick (kg)",
        'dim_label': "Dimension limit for 1-by-1 (cm)",
        'hmat_label': "Max pieces per grab (light parts)",
        'exclude_label': "Exclude materials:",
        'sec_ratio': "🎯 Data Reliability & Calculation Source",
        'ratio_desc': "Data foundation used for calculating physical movements:",
        'ratio_master': "Exact (Boxes / Pallets)",
        'ratio_loose_ok': "Exact (Verified Loose)",
        'ratio_loose_miss': "Estimated (Missing Box)",
        'exp_missing_data': "🔍 Show materials with missing box data (Estimates Leaderboard)",
        'sec_queue_title': "📊 Average Workload by Queue",
        'q_col_queue': "Queue",
        'q_col_to': "TOs",
        'q_col_orders': "Orders",
        'q_col_loc': "Avg Locs",
        'q_col_pcs': "Avg Pieces",
        'q_col_moves': "Avg Moves per TO",
        'q_col_box': "Avg Boxes",
        'q_pct_box': "% Boxes",
        'q_col_ok': "Avg Loose",
        'q_pct_ok': "% Loose",
        'q_col_miss': "Avg Missing",
        'q_pct_miss': "% Missing",
        'sec_queue_top_title': "🏆 TOP 100 Materials by Queue",
        'q_select': "Show TOP 100 for Queue:",
        'sec1_title': "🎯 Single-Material Pallet Orders",
        'm_orders': "Orders",
        'm_qty': "Avg Pcs / Order",
        'm_pos': "Avg Bins / Order",
        'm_mov': "Avg Physical Moves",
        'exp_detail_title': "Show Orders Table (1 Material)",
        'col_mat': "Material",
        'col_qty': "Total Pieces",
        'col_mov': "Total Movements",
        'col_mov_box': "Moves (Boxes)",
        'col_mov_loose_ok': "Moves (Loose)",
        'col_mov_loose_miss': "Moves (Missing Box)",
        'col_wgt': "Weight (kg)",
        'col_max_dim': "Max Dim (cm)",
        'sec1_top_title': "🏆 TOP 100 Materials for Pallet Orders",
        'btn_download': "📥 Download Comprehensive Report (Excel)",
        'no_orders': "No orders found.",
        'sec3_title': "🔍 Master Data Viewer",
        'search_label': "Check specific material data:",
        'tab_dashboard': "📊 Dashboard & Queue",
        'tab_pallets': "📦 Pallet Orders",
        'tab_top': "🏆 TOP Materials & Data Quality",
        'tab_audit': "🔍 Tools & Audit",
        'col_lines': "Lines",
        'col_box': "Packaging",
        'val_loose': "Loose"
    }
}

def t(key): return TEXTS[st.session_state.lang][key]

def get_match_key(val):
    v = str(val).strip().upper()
    if '.' in v and v.replace('.', '').isdigit():
        return v.rstrip('0').rstrip('.')
    return v

def create_top_100_df(df_subset):
    if df_subset is None or df_subset.empty: return pd.DataFrame()
    agg = df_subset.groupby('Material').agg(
        pocet_picku=('Material', 'count'), celkove_mnozstvi=('Qty', 'sum'),
        celkem_pohybu=('Pohyby_Rukou', 'sum'), pohyby_box=('Pohyby_Box', 'sum'),
        pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'), pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
        celkova_natacena_vaha=('Celkova_Vaha_KG', 'sum'), Box_Sizes_List=('Box_Sizes_List', 'first')
    ).reset_index()

    agg.rename(columns={
        'Material': t('col_mat'), 'pocet_picku': t('col_lines'),
        'celkem_pohybu': t('col_mov'), 'pohyby_box': t('col_mov_box'),
        'pohyby_loose_ok': t('col_mov_loose_ok'), 'pohyby_loose_miss': t('col_mov_loose_miss'),
        'celkove_mnozstvi': t('col_qty'), 'celkova_natacena_vaha': t('col_wgt')
    }, inplace=True)

    return agg.sort_values(by=t('col_mov'), ascending=False).head(100)[[t('col_mat'), t('col_lines'), t('col_qty'), t('col_wgt'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_mov')]]

def add_excel_chart(writer, sheet_name, df_top):
    worksheet = writer.sheets[sheet_name]
    chart = BarChart()
    chart.type, chart.style, chart.width, chart.height = "col", 10, 25, 12
    chart.title = "Zátěž materiálů dle fyzických pohybů"
    col_mat_idx = list(df_top.columns).index(t('col_mat')) + 1
    col_mov_idx = list(df_top.columns).index(t('col_mov')) + 1
    chart.add_data(Reference(worksheet, min_col=col_mov_idx, min_row=1, max_row=len(df_top)+1), titles_from_data=True)
    chart.set_categories(Reference(worksheet, min_col=col_mat_idx, min_row=2, max_row=len(df_top)+1))
    chart.legend = None 
    worksheet.add_chart(chart, "K2")

def main():
    col_title, col_lang = st.columns([8, 1])
    with col_title:
        st.title(t('title'))
        st.markdown(f"*{t('desc')}*")
    with col_lang:
        if st.button(t('switch_lang')):
            st.session_state.lang = 'en' if st.session_state.lang == 'cs' else 'cs'
            st.rerun()

    st.divider()

    st.sidebar.header(t('sidebar_title'))
    limit_vahy = st.sidebar.number_input(t('weight_label'), min_value=0.1, max_value=20.0, value=2.0, step=0.5)
    limit_rozmeru = st.sidebar.number_input(t('dim_label'), min_value=1.0, max_value=200.0, value=15.0, step=1.0)
    kusy_na_hmat = st.sidebar.slider(t('hmat_label'), min_value=1, max_value=20, value=3, step=1)
    
    with st.expander(t('upload_title'), expanded=True):
        uploaded_files = st.file_uploader(t('upload_help'), type=['csv', 'xlsx'], accept_multiple_files=True)

    if uploaded_files:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        df_pick, df_marm, df_manual, df_queue = None, None, None, None

        status_text.markdown("**🔄 Načítání a čtení vstupních souborů (20 %)...**")
        progress_bar.progress(20)
        time.sleep(0.3)

        for file in uploaded_files:
            temp_df = pd.read_csv(file, dtype=str) if file.name.lower().endswith('.csv') else pd.read_excel(file, dtype=str)
            if 'Delivery' in temp_df.columns and 'Act.qty (dest)' in temp_df.columns: df_pick = temp_df
            elif 'Numerator' in temp_df.columns and 'Alternative Unit of Measure' in temp_df.columns: df_marm = temp_df
            elif 'Queue' in temp_df.columns and ('Transfer Order Number' in temp_df.columns or 'SD Document' in temp_df.columns): df_queue = temp_df
            elif len(temp_df.columns) >= 2: df_manual = temp_df

        if df_pick is None:
            st.error(t('err_pick'))
            progress_bar.empty()
            status_text.empty()
            return

        status_text.markdown("**⚙️ Zpracování Master Dat a ručních ověření (40 %)...**")
        progress_bar.progress(40)
        
        df_pick['Material'] = df_pick['Material'].astype(str).str.strip()
        df_pick['Match_Key'] = df_pick['Material'].apply(get_match_key)
        
        users_to_remove = ['UIDJ5089', 'UIH25501']
        mask_admins = pd.Series(False, index=df_pick.index)
        for col in df_pick.columns: mask_admins = mask_admins | df_pick[col].isin(users_to_remove)
        num_removed_admins = mask_admins.sum()
        if num_removed_admins > 0: df_pick = df_pick[~mask_admins].copy()
        
        df_pick = df_pick.dropna(subset=['Delivery', 'Material']).copy()
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)
        df_pick['Source Storage Bin'] = df_pick.get('Source Storage Bin', df_pick.get('Storage Bin', ''))

        status_text.markdown("**🔗 Párování lokací, zakázek a Queue (60 %)...**")
        progress_bar.progress(60)

        queue_count_col = 'Delivery'
        if df_queue is not None:
            if 'Transfer Order Number' in df_pick.columns and 'Transfer Order Number' in df_queue.columns:
                q_map = df_queue.dropna(subset=['Transfer Order Number', 'Queue']).drop_duplicates('Transfer Order Number').set_index('Transfer Order Number')['Queue'].to_dict()
                df_pick['Queue'] = df_pick['Transfer Order Number'].map(q_map)
                queue_count_col = 'Transfer Order Number'
                for d_col in ['Confirmation Date', 'Creation Date']:
                    if d_col in df_queue.columns:
                        df_pick['Date'] = df_pick['Transfer Order Number'].map(df_queue.dropna(subset=['Transfer Order Number', d_col]).drop_duplicates('Transfer Order Number').set_index('Transfer Order Number')[d_col].to_dict())
                        break
            elif 'SD Document' in df_queue.columns:
                q_map = df_queue.dropna(subset=['SD Document', 'Queue']).drop_duplicates('SD Document').set_index('SD Document')['Queue'].to_dict()
                df_pick['Queue'] = df_pick['Delivery'].map(q_map)
                for d_col in ['Confirmation Date', 'Creation Date']:
                    if d_col in df_queue.columns:
                        df_pick['Date'] = df_pick['Delivery'].map(df_queue.dropna(subset=['SD Document', d_col]).drop_duplicates('SD Document').set_index('SD Document')[d_col].to_dict())
                        break
            if 'Queue' in df_pick.columns: df_pick = df_pick[df_pick['Queue'].astype(str).str.upper() != 'CLEARANCE'].copy()
        else:
            df_pick['Queue'], df_pick['Date'] = 'N/A', np.nan

        df_pick['Month'] = pd.to_datetime(df_pick.get('Date', np.nan), errors='coerce').dt.to_period('M').astype(str).replace('NaT', 'Neznámé')
        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().str.upper()
        zakazky_s_x = df_pick[df_pick['Removal of total SU'] == 'X']['Delivery'].unique()

        manual_boxes = {}
        if df_manual is not None and not df_manual.empty:
            c_mat, c_pkg = df_manual.columns[0], df_manual.columns[1]
            for _, row in df_manual.iterrows():
                if pd.isna(row[c_mat]) or str(row[c_mat]).upper() in ['NAN', 'NONE', '']: continue
                mat_key, pkg = get_match_key(str(row[c_mat])), str(row[c_pkg])
                nums = re.findall(r'(\d+)\s*(?:ks|kus|pcs)|\bK-(\d+)\b|(?:pytl[íi]k|pytel|role|balen[íi]|krabice|karton|box)[^\d]*(\d+)', pkg, flags=re.IGNORECASE)
                ext = sorted(list(set([int(g) for m in nums for g in m if g])), reverse=True)
                if not ext and 'po kusech' in pkg.lower(): ext = [1]
                if ext: manual_boxes[mat_key] = ext

        box_dict, weight_dict, dim_dict = {}, {}, {}
        if df_marm is not None:
            df_marm['Match_Key'] = df_marm['Material'].apply(get_match_key)
            df_boxes = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK', 'KAR', 'KART', 'PAK', 'VPE', 'CAR', 'BLO'])].copy()
            df_boxes['Numerator'] = pd.to_numeric(df_boxes['Numerator'], errors='coerce').fillna(0)
            box_dict = df_boxes.groupby('Match_Key')['Numerator'].apply(lambda g: sorted([int(x) for x in g if x > 1], reverse=True)).to_dict()

            df_st = df_marm[df_marm['Alternative Unit of Measure'].isin(['ST', 'PCE', 'KS'])].copy()
            df_st['Gross Weight'] = pd.to_numeric(df_st['Gross Weight'], errors='coerce').fillna(0)
            df_st['Weight_KG'] = np.where(df_st['Unit of Weight'].astype(str).str.upper() == 'G', df_st['Gross Weight']/1000.0, df_st['Gross Weight'])
            weight_dict = df_st.groupby('Match_Key')['Weight_KG'].first().to_dict()

            def to_cm(val, unit):
                try:
                    v, u = float(val), str(unit).upper().strip()
                    if u == 'MM': return v / 10.0
                    if u == 'M': return v * 100.0
                    return v 
                except: return 0.0

            for dim in ['Length', 'Width', 'Height']: df_st[dim[0]] = df_st.apply(lambda r: to_cm(r[dim], r['Unit of Dimension']), axis=1)
            dim_dict = df_st.set_index('Match_Key')[['L', 'W', 'H']].max(axis=1).to_dict()

        df_pick['Box_Sizes_List'] = df_pick['Match_Key'].apply(lambda m: manual_boxes.get(m, box_dict.get(m, [])))
        df_pick['Piece_Weight_KG'] = df_pick['Match_Key'].map(weight_dict).fillna(0)
        df_pick['Piece_Max_Dim_CM'] = df_pick['Match_Key'].map(dim_dict).fillna(0)

        excluded_materials = st.sidebar.multiselect(t('exclude_label'), options=sorted(df_pick['Material'].unique()), default=[])
        if excluded_materials: df_pick = df_pick[~df_pick['Material'].isin(excluded_materials)]

        status_text.markdown("**🤖 Simulace fyzických pohybů a ergonomie (85 %)...**")
        progress_bar.progress(85)
        time.sleep(0.3)

        def spocitej_pohyby_detail(row):
            qty = row['Qty']
            if qty <= 0: return 0, 0, 0, 0
            if str(row.get('Removal of total SU', '')).strip().upper() == 'X': return 1, 1, 0, 0
            
            pb, pok, pmiss, zbytek = 0, 0, 0, qty
            for b in row['Box_Sizes_List']:
                if b > 1 and zbytek >= b:
                    pb += zbytek // b
                    zbytek %= b
            if zbytek > 0:
                p = zbytek if (row['Piece_Weight_KG'] >= limit_vahy or row['Piece_Max_Dim_CM'] >= limit_rozmeru) else np.ceil(zbytek / kusy_na_hmat)
                if len(row['Box_Sizes_List']) == 0: pmiss += p
                else: pok += p
            return pb + pok + pmiss, pb, pok, pmiss

        df_pick[['Pohyby_Rukou', 'Pohyby_Box', 'Pohyby_Loose_OK', 'Pohyby_Loose_Miss']] = df_pick.apply(spocitej_pohyby_detail, axis=1, result_type='expand')
        df_pick['Celkova_Vaha_KG'] = df_pick['Qty'] * df_pick['Piece_Weight_KG']

        status_text.markdown("**✅ Hotovo! Sestavuji Dashboardy (100 %)**")
        progress_bar.progress(100)
        time.sleep(0.5)
        progress_bar.empty()
        status_text.empty()

        # ==========================================
        # 3. ZPRACOVÁNÍ VÝSLEDKŮ DO TABS
        # ==========================================
        tab_dash, tab_pallets, tab_top, tab_audit = st.tabs([t('tab_dashboard'), t('tab_pallets'), t('tab_top'), t('tab_audit')])

        with st.container():
            col_i1, col_i2, col_i3 = st.columns(3)
            if num_removed_admins > 0: col_i1.info(t('info_users').format(num_removed_admins))
            if len(zakazky_s_x) > 0: col_i2.warning(t('info_clean').format(len(zakazky_s_x)))
            if manual_boxes: col_i3.success(t('info_manual').format(len(manual_boxes)))

        # === TAB 1: DASHBOARD A QUEUE ===
        with tab_dash:
            tot_mov = df_pick['Pohyby_Rukou'].sum()
            if tot_mov > 0:
                st.subheader(t('sec_ratio'))
                st.write(t('ratio_desc'))
                c_r1, c_r2, c_r3 = st.columns(3)
                c_r1.metric(t('ratio_master'), f"{(df_pick['Pohyby_Box'].sum() / tot_mov * 100):.1f} %", f"{df_pick['Pohyby_Box'].sum():,.0f} pohybů".replace(',', ' '))
                c_r2.metric(t('ratio_loose_ok'), f"{(df_pick['Pohyby_Loose_OK'].sum() / tot_mov * 100):.1f} %", f"{df_pick['Pohyby_Loose_OK'].sum():,.0f} pohybů".replace(',', ' '))
                c_r3.metric(t('ratio_loose_miss'), f"{(df_pick['Pohyby_Loose_Miss'].sum() / tot_mov * 100):.1f} %", f"{df_pick['Pohyby_Loose_Miss'].sum():,.0f} pohybů".replace(',', ' '), delta_color="inverse")

            queue_summary = None
            if 'Queue' in df_pick.columns and df_pick['Queue'].notna().any() and df_pick['Queue'].nunique() > 1:
                st.divider()
                st.subheader(t('sec_queue_title'))
                
                sel_month = st.selectbox("📅 Filtrovat podle měsíce:", options=['Všechny měsíce'] + sorted([m for m in df_pick['Month'].unique() if m != 'Neznámé']) + (['Neznámé'] if 'Neznámé' in df_pick['Month'].unique() else []))
                df_q_filter = df_pick[df_pick['Month'] == sel_month] if sel_month != 'Všechny měsíce' else df_pick.copy()

                if not df_q_filter.empty:
                    queue_agg_raw = df_q_filter.groupby([queue_count_col, 'Queue']).agg(
                        celkem_pohybu=('Pohyby_Rukou', 'sum'), pohyby_box=('Pohyby_Box', 'sum'),
                        pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'), pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
                        total_qty=('Qty', 'sum'), num_materials=('Material', 'nunique'),
                        pocet_lokaci=('Source Storage Bin', 'nunique'), delivery=('Delivery', 'first')
                    ).reset_index()
                    
                    def adjust_queue_name(row):
                        return row['Queue'] + (' (Single)' if row['num_materials'] == 1 else ' (Mix)') if str(row['Queue']).upper() in ['PI_PL', 'PI_PL_OE'] else row['Queue']

                    totals_rows = queue_agg_raw[queue_agg_raw['Queue'].str.upper().isin(['PI_PL', 'PI_PL_OE'])].copy()
                    totals_rows['Queue'] += ' (Total)'
                    queue_agg_raw['Queue'] = queue_agg_raw.apply(adjust_queue_name, axis=1)
                    queue_agg_final = pd.concat([queue_agg_raw, totals_rows], ignore_index=True)
                    
                    q_sum = queue_agg_final.groupby('Queue').agg(
                        pocet_zakazek=('delivery', 'nunique'), prum_lokaci=('pocet_lokaci', 'mean'),
                        prum_kusu=('total_qty', 'mean'), prum_pohybu=('celkem_pohybu', 'mean'),
                        prum_box=('pohyby_box', 'mean'), prum_ok=('pohyby_loose_ok', 'mean'), prum_miss=('pohyby_loose_miss', 'mean')
                    )
                    q_sum['pocet_TO'] = queue_agg_final.groupby('Queue')[queue_count_col].nunique() if queue_count_col == 'Transfer Order Number' else q_sum['pocet_zakazek']
                    q_sum = q_sum.reset_index().sort_values('prum_pohybu', ascending=False)
                    
                    for k in ['box', 'ok', 'miss']: q_sum[f'pct_{k}'] = np.where(q_sum['prum_pohybu'] > 0, (q_sum[f'prum_{k}'] / q_sum['prum_pohybu']) * 100, 0)
                    
                    display_q = q_sum[['Queue', 'pocet_TO', 'pocet_zakazek', 'prum_lokaci', 'prum_kusu', 'prum_pohybu', 'prum_box', 'pct_box', 'prum_ok', 'pct_ok', 'prum_miss', 'pct_miss']].copy()
                    display_q.columns = [t('q_col_queue'), t('q_col_to'), t('q_col_orders'), t('q_col_loc'), t('q_col_pcs'), t('q_col_moves'), t('q_col_box'), t('q_pct_box'), t('q_col_ok'), t('q_pct_ok'), t('q_col_miss'), t('q_pct_miss')]
                    
                    col_qt1, col_qt2 = st.columns([2.5, 1])
                    with col_qt1:
                        st.dataframe(display_q.style.format({c: "{:.1f}" for c in display_q.columns if 'Prům' in c} | {c: "{:.1f} %" for c in display_q.columns if '%' in c}), use_container_width=True, hide_index=True)
                    with col_qt2:
                        st.bar_chart(q_sum.set_index('Queue')['prum_pohybu'])

        # === TAB 2: PALETOVÉ ZAKÁZKY ===
        with tab_pallets:
            st.subheader(t('sec1_title'))
            grouped_orders = df_pick.groupby('Delivery').agg(
                num_materials=('Material', 'nunique'), material=('Material', 'first'),
                certs=('Certificate Number', lambda x: x.dropna().unique().tolist()),
                total_qty=('Qty', 'sum'), num_positions=('Source Storage Bin', 'nunique'),
                celkem_pohybu=('Pohyby_Rukou', 'sum'), pohyby_box=('Pohyby_Box', 'sum'), 
                pohyby_loose_ok=('Pohyby_Loose_OK', 'sum'), pohyby_loose_miss=('Pohyby_Loose_Miss', 'sum'),
                vaha_zakazky=('Celkova_Vaha_KG', 'sum'), max_rozmer=('Piece_Max_Dim_CM', 'first')
            )
            filtered_orders = grouped_orders[(~grouped_orders.index.isin(zakazky_s_x)) & (grouped_orders['num_materials'] == 1) & (grouped_orders['certs'].apply(lambda c: len([x for x in c if pd.notna(x) and str(x).strip() and not str(x).strip().startswith('460')]) > 0))]

            if not filtered_orders.empty:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric(t('m_orders'), f"{len(filtered_orders):,}".replace(',', ' '))
                c2.metric(t('m_qty'), f"{filtered_orders['total_qty'].mean():.1f}")
                c3.metric(t('m_pos'), f"{filtered_orders['num_positions'].mean():.2f}")
                c4.metric(t('m_mov'), f"{filtered_orders['celkem_pohybu'].mean():.1f}")

                tot_p_pal = filtered_orders['celkem_pohybu'].sum()
                if tot_p_pal > 0:
                    st.markdown("**Spolehlivost dat čistě pro tyto paletové zakázky:**")
                    c_p1, c_p2, c_p3 = st.columns(3)
                    c_p1.metric(t('ratio_master'), f"{(filtered_orders['pohyby_box'].sum() / tot_p_pal * 100):.1f} %")
                    c_p2.metric(t('ratio_loose_ok'), f"{(filtered_orders['pohyby_loose_ok'].sum() / tot_p_pal * 100):.1f} %")
                    c_p3.metric(t('ratio_loose_miss'), f"{(filtered_orders['pohyby_loose_miss'].sum() / tot_p_pal * 100):.1f} %", delta_color="inverse")

                with st.expander(t('exp_detail_title')):
                    display_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'pohyby_box', 'pohyby_loose_ok', 'pohyby_loose_miss', 'vaha_zakazky', 'max_rozmer', 'certs']].copy()
                    display_df.columns = [t('col_mat'), t('col_qty'), t('col_mov'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_wgt'), t('col_max_dim'), 'Certifikát']
                    st.dataframe(display_df, use_container_width=True)
            else:
                st.warning(t('no_orders'))

        # === TAB 3: TOP MATERIÁLY ===
        with tab_top:
            st.subheader(t('sec_queue_top_title'))
            if 'Queue' in df_pick.columns and df_pick['Queue'].notna().any() and df_pick['Queue'].nunique() > 1 and not df_q_filter.empty:
                selected_queue_disp = st.selectbox(t('q_select'), options=sorted(queue_agg_final['Queue'].dropna().unique().tolist()))
                if '(Total)' in selected_queue_disp: df_queue_subset = df_q_filter[df_q_filter['Queue'] == selected_queue_disp.replace(' (Total)', '')]
                elif '(Single)' in selected_queue_disp: df_queue_subset = df_q_filter[(df_q_filter['Queue'] == selected_queue_disp.replace(' (Single)', '')) & (df_q_filter[queue_count_col].isin(queue_agg_raw[(queue_agg_raw['Queue'] == selected_queue_disp) & (queue_agg_raw['num_materials'] == 1)][queue_count_col]))]
                elif '(Mix)' in selected_queue_disp: df_queue_subset = df_q_filter[(df_q_filter['Queue'] == selected_queue_disp.replace(' (Mix)', '')) & (df_q_filter[queue_count_col].isin(queue_agg_raw[(queue_agg_raw['Queue'] == selected_queue_disp) & (queue_agg_raw['num_materials'] > 1)][queue_count_col]))]
                else: df_queue_subset = df_q_filter[df_q_filter['Queue'] == selected_queue_disp]

                top_100_queue = create_top_100_df(df_queue_subset)
                if not top_100_queue.empty:
                    col_q1, col_q2 = st.columns([1.5, 1])
                    with col_q1: st.dataframe(top_100_queue.style.format({t('col_wgt'): "{:.1f}"} | {c: "{:.0f}" for c in top_100_queue.columns if 'Pohyby' in c}), use_container_width=True, hide_index=True)
                    with col_q2: st.bar_chart(top_100_queue.set_index(t('col_mat'))[t('col_mov')])

            st.divider()
            st.subheader(t('exp_missing_data').replace('🔍 ', ''))
            all_mat_agg = df_pick.groupby('Material').agg(lines=('Material', 'count'), qty=('Qty', 'sum'), miss=('Pohyby_Loose_Miss', 'sum'), mov=('Pohyby_Rukou', 'sum')).reset_index()
            all_mat_agg.columns = [t('col_mat'), t('col_lines'), t('col_qty'), t('col_mov_loose_miss'), t('col_mov')]
            miss_df = all_mat_agg[all_mat_agg[t('col_mov_loose_miss')] > 0].sort_values(by=t('col_mov_loose_miss'), ascending=False).head(100)
            if not miss_df.empty: st.dataframe(miss_df.style.format({c: "{:.0f}" for c in [t('col_mov_loose_miss'), t('col_mov')]}), use_container_width=True, hide_index=True)
            else: st.success("Všechna data o baleních jsou k dispozici, žádné odhady!")

        # === TAB 4: NÁSTROJE A AUDIT ===
        with tab_audit:
            # 5x Namátková kontrola pro každou Queue
            st.subheader("🎲 Detailní Audit logiky (5 úkolů za každou frontu)")
            st.write("Slouží pro obhajobu výpočtů s klientem. Vygeneruje až 5 náhodných úkolů z každé existující fronty s maximálním detailem.")
            
            if st.button("Vygenerovat auditní report (5 úkolů z každé fronty)", type="primary"):
                if len(df_pick) > 0:
                    audit_samples = {}
                    # Vyfiltrujeme jen reálné fronty, ze kterých máme data (ignorujeme N/A pokud jich je tam málo)
                    valid_queues = sorted([q for q in df_pick['Queue'].dropna().unique() if q != 'N/A'])
                    
                    for q in valid_queues:
                        q_data = df_pick[df_pick['Queue'] == q]
                        unique_tos = q_data[queue_count_col].dropna().unique()
                        if len(unique_tos) > 0:
                            # Vybere náhodně maximálně 5 unikátních TO (pokud jich je méně než 5, vezme všechny)
                            sampled = np.random.choice(unique_tos, min(5, len(unique_tos)), replace=False)
                            audit_samples[q] = sampled
                            
                    st.session_state['audit_samples'] = audit_samples

            if 'audit_samples' in st.session_state:
                for q, tos in st.session_state['audit_samples'].items():
                    # Pro každou Queue vytvoříme samostatný panel
                    with st.expander(f"📁 Fronta: {q} (Zobrazeno {len(tos)} úkolů)", expanded=False):
                        for i, r_to in enumerate(tos, 1):
                            st.markdown(f"#### {i}. Úkol (TO / Doklad): **`{r_to}`**")
                            to_data = df_pick[df_pick[queue_count_col] == r_to]
                            
                            deliv_val = to_data['Delivery'].iloc[0] if 'Delivery' in to_data.columns else 'Neznámé'
                            date_val = to_data['Date'].iloc[0] if 'Date' in to_data.columns else 'Neznámé'
                            st.caption(f"**Zakázka (Delivery):** `{deliv_val}` | **Zpracováno dne:** `{date_val}`")
                            
                            for _, row in to_data.iterrows():
                                mat = row['Material']
                                qty = row['Qty']
                                boxes = row.get('Box_Sizes_List', [])
                                w = row.get('Piece_Weight_KG', 0)
                                d = row.get('Piece_Max_Dim_CM', 0)
                                su = row.get('Removal of total SU', '')
                                src_bin = row.get('Source Storage Bin', 'Neznámá')
                                
                                st.markdown(f"**Materiál:** `{mat}` | **Zdrojová lokace:** `{src_bin}` | **Množství:** {qty} ks | **Balení (Master data):** {boxes if boxes else 'Chybí'} | **Váha/ks:** {w:.3f} kg | **Max rozměr:** {d:.1f} cm")
                                
                                if su == 'X':
                                    st.info(f"➡️ Z lokace `{src_bin}` byla odebrána celá manipulační jednotka (paleta) označená jako 'X'. -> **Započítán 1 pohyb.**")
                                else:
                                    zbytek = qty
                                    if boxes:
                                        for b in boxes:
                                            if b > 1 and zbytek >= b:
                                                m = zbytek // b
                                                st.info(f"➡️ Z lokace `{src_bin}` odebráno **{m} krabic** (po {b} ks) = **{m} pohybů**. (Zbylo {zbytek % b} ks)")
                                                zbytek %= b
                                    if zbytek > 0:
                                        if w >= limit_vahy or d >= limit_rozmeru:
                                            st.warning(f"➡️ Zbylých {zbytek} ks překračuje limity pro hrst ({w:.3f}kg, {d:.1f}cm). Musí se brát po jednom kuse -> **{zbytek} pohybů**.")
                                        else:
                                            hmaty = np.ceil(zbytek / kusy_na_hmat)
                                            st.success(f"➡️ Zbylých {zbytek} ks je drobných. Lze je brát do hrsti (max {kusy_na_hmat} ks najednou) -> **{hmaty:.0f} pohybů**.")
                                
                                st.markdown(f"> **Započítáno fyzických pohybů pro tento řádek:** `{row.get('Pohyby_Rukou', 0)}`")
                                st.write("---")

            # Vyhledávač materiálů
            st.divider()
            st.subheader(t('sec3_title'))
            mat_search = st.selectbox(t('search_label'), options=[""] + sorted(df_pick['Material'].unique().tolist()))
            if mat_search:
                search_key = get_match_key(mat_search)
                if search_key in manual_boxes:
                    if manual_boxes[search_key] == [1]: st.success("✅ **Ruční ověření:** Nastaveno natvrdo jako **Volné kusy (1 ks)**.")
                    else: st.success(f"✅ **Ruční ověření nalezeno:** Nastaveny krabice/pytlíky po **{manual_boxes[search_key]} ks**.")
                else: st.info("ℹ️ Tento materiál nemá zadané žádné ruční ověření.")
                
                c_info1, c_info2 = st.columns(2)
                c_info1.metric("Váha 1 kusu (z MARM)", f"{weight_dict.get(search_key, 0):.3f} kg")
                c_info2.metric("Nejdelší rozměr (z MARM)", f"{dim_dict.get(search_key, 0):.1f} cm")
                
                if df_marm is not None:
                    st.write("**Surová data z MARM reportu:**")
                    marm_detail = df_marm[df_marm['Match_Key'] == search_key]
                    if not marm_detail.empty: st.dataframe(marm_detail[['Alternative Unit of Measure', 'Numerator', 'Denominator', 'Gross Weight', 'Unit of Weight', 'Length', 'Width', 'Height', 'Unit of Dimension']], hide_index=True, use_container_width=True)

           # Export
            st.divider()
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # 1. Metodika a Nastavení
                pd.DataFrame({"Téma": ["Hranice váhy", "Hranice rozměru", "Max do hrsti"], "Nastavení": [f"{limit_vahy} kg", f"{limit_rozmeru} cm", f"{kusy_na_hmat} ks"]}).to_excel(writer, index=False, sheet_name='Info_a_Metodika')
                
                # 2. Agregované přehledy pro rychlé prohlížení
                if 'display_q' in locals(): display_q.to_excel(writer, index=False, sheet_name='Analyza_Queue')
                if 'filtered_orders' in locals() and not filtered_orders.empty:
                    ex_df = filtered_orders[['material', 'total_qty', 'celkem_pohybu', 'pohyby_box', 'pohyby_loose_ok', 'pohyby_loose_miss', 'vaha_zakazky', 'max_rozmer']].copy()
                    ex_df.columns = [t('col_mat'), t('col_qty'), t('col_mov'), t('col_mov_box'), t('col_mov_loose_ok'), t('col_mov_loose_miss'), t('col_wgt'), t('col_max_dim')]
                    ex_df.to_excel(writer, index=True, sheet_name='Souhrn_Zakazek_1_Mat')
                df_pick.groupby('Material').agg(Pohyby=('Pohyby_Rukou', 'sum'), Kusy=('Qty', 'sum')).reset_index().to_excel(writer, index=False, sheet_name='Vsechna_Data_Agregovano')

                # 3. ZDROJOVÁ TABULKA FAKTŮ PRO POWER BI
                pbi_df = df_pick.copy()
                # Excel neumí ukládat datový typ "seznam/list", proto strukturu balení převedeme na čitelný text
                pbi_df['Box_Sizes_List'] = pbi_df['Box_Sizes_List'].astype(str)
                pbi_df.to_excel(writer, index=False, sheet_name='PowerBI_Data')

            st.download_button(label="📥 Stáhnout data (Připraveno pro Power BI)", data=buffer.getvalue(), file_name="Skladova_Analytika.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

if __name__ == "__main__":
    main()
