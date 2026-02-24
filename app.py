import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import time
from openpyxl.chart import BarChart, Reference

# ==========================================
# 1. NASTAVENÍ STRÁNKY A CSS STYLING
# ==========================================
st.set_page_config(
    page_title="Analýza pickování", 
    page_icon="📦", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Custom CSS pro vzhled karet a grafů
st.markdown("""
    <style>
    div[data-testid="metric-container"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        padding: 5% 5% 5% 10%;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

if 'lang' not in st.session_state:
    st.session_state.lang = 'cs'

# ==========================================
# 2. SLOVNÍKY A LOKALIZACE (PLNÁ VERZE)
# ==========================================
QUEUE_DESC = {
    'PI_PL (Mix)': 'Mix Pallet',
    'PI_PL (Total)': 'Mix Pallet',
    'PI_PL (Single)': 'Mix Pallet',
    'PI_PL_OE (Mix)': 'Mix Pallet OE',
    'PI_PA_OE': 'Parcel OE',
    'PI_PL_OE (Total)': 'Mix Pallet OE',
    'PI_PL_OE (Single)': 'Mix Pallet OE',
    'PI_PA': 'Parcel',
    'PI_PA_RU': 'Parcel Express',
    'PI_PL_FU': 'Full Pallet',
    'PI_PL_FUOE': 'Full Pallet OE'
}

TEXTS = {
    'cs': {
        'switch_lang': "🇬🇧 Switch to English",
        'title': "📦 Analýza pickování",
        'desc': "Nástroj pro modelování fyzické zátěže pickování",
        'upload_title': "📁 Nahrání vstupních dat (Klikněte pro sbalení/rozbalení)",
        'upload_help': "Nahrajte Pick report, MARM report, TO details (Queue) a volitelně i ruční ověření balení.",
        'info_users': "💡 Vyloučeno **{} systémových řádků** (UIDJ5089, UIH25501).",
        'info_clean': "💡 Započítán 1 pohyb pro **{} řádků** 'X' (Platí POUZE pro Queue: PI_PL_FU, PI_PL_FUOE).",
        'info_manual': "✅ Načteno ruční ověření pro **{} unikátních materiálů**.",
        'sidebar_title': "⚙️ Konfigurace algoritmů",
        'weight_label': "Hranice pro nošení po 1 ks (kg)",
        'dim_label': "Hranice rozměru pro 1 ks (cm)",
        'hmat_label': "Max ks lehkých dílů do hrsti",
        'exclude_label': "Vyloučit materiály z výpočtů:",
        'sec_ratio': "🎯 Spolehlivost dat a zdroj výpočtů",
        'ratio_desc': "Z jakých podkladů aplikace vycházela (Ukazatel kvality dat ze SAPu):",
        'ratio_moves': "Podíl z celkového počtu POHYBŮ:",
        'ratio_exact': "Přesně (Známe balení)",
        'ratio_miss': "Odhady (Chybí balení)",
        'logic_explain_title': "ℹ️ Jak aplikace počítá fyzické pohyby?",
        'logic_explain_text': """Tento algoritmus simuluje reálnou lidskou námahu ve skladu:
        1. **Plné jednotky (X):** Pokud je ve frontách Full Pallet (FU/FUOE) označen odběr 'X', započítá se 1 pohyb (manipulace ještěrkou).
        2. **Krabice (Exact):** U ostatních front algoritmus prioritně hledá celá balení (krabice) dle Master dat. Každá krabice = 1 pohyb.
        3. **Volné kusy (Exact):** Pokud známe balení, zbylé kusy dělíme počtem 'do hrsti'. Pokud je kus těžký nebo velký, bere se po 1 ks.
        4. **Odhady (Miss):** Pokud SAP nezná rozměr krabice, aplikace aplikuje bezpečností odhad zátěže, aby nedocházelo k podhodnocení námahy.""",
        'exp_missing_data': "🔍 Zobrazit materiály s chybějícími daty o balení (Žebříček 'odhadů')",
        'sec_queue_title': "📊 Průměrná náročnost dle typu pickování (Queue)",
        'filter_month': "📅 Filtrovat podle měsíce:",
        'all_months': "Všechny měsíce",
        'all_queues': "Všechny Queue dohromady",
        'unknown': "Neznámé",
        'q_col_queue': "Queue",
        'q_col_desc': "Popis",
        'q_col_to': "Počet TO",
        'q_col_orders': "Zakázky",
        'q_col_loc': "Prům. lokací",
        'q_col_pcs': "Prům. kusů",
        'q_col_mov_loc': "Pohybů na lokaci",
        'q_col_exact_loc': "Přesně na lokaci",
        'q_pct_exact': "% Přesně",
        'q_col_miss_loc': "Odhad na lokaci",
        'q_pct_miss': "% Odhad",
        'sec_queue_top_title': "🏆 TOP 100 materiálů podle Queue",
        'q_select': "Zobrazit TOP 100 pro:",
        'sec1_title': "🎯 Analýza paletových zakázek (Mix Pallet)",
        'pallets_clean_info': "*(Počítáno pouze z front PI_PL a PI_PL_OE)*",
        'm_orders': "Počet zakázek",
        'm_qty': "Prům. kusů / zakázku",
        'm_pos': "Prům. pozic / zakázku",
        'm_mov_loc': "Prům. fyz. pohybů na lokaci",
        'exp_detail_title': "Zobrazit tabulku zakázek",
        'col_mat': "Materiál",
        'col_qty': "Kusů celkem",
        'col_mov': "Celkem pohybů",
        'col_mov_exact': "Pohyby (Přesně)",
        'col_mov_miss': "Pohyby (Odhady)",
        'col_wgt': "Hmotnost (kg)",
        'col_max_dim': "Rozměr (cm)",
        'audit_title': "🎲 Detailní Auditní Report (Náhodné vzorky)",
        'audit_phys_moves': "Fyzických pohybů",
        'audit_gen_btn': "Vygenerovat náhodný Audit (5 úkolů z každé fronty)",
        'sec3_title': "🔍 Prohlížeč Master Dat",
        'search_label': "Zkontrolujte si konkrétní materiál:",
        'tab_dashboard': "📊 Dashboard & Queue",
        'tab_pallets': "📦 Paletové zakázky",
        'tab_top': "🏆 TOP Materiály",
        'tab_audit': "🔍 Nástroje & Audit",
        'col_lines': "Řádky",
        'btn_download': "📥 Stáhnout kompletní report (Excel)",
        'err_pick': "Chyba: Pick report nebyl nalezen ve vstupech."
    },
    'en': {
        'switch_lang': "🇨🇿 Přepnout do češtiny",
        'title': "📦 Picking Analysis",
        'desc': "Tool for modeling physical picking workload",
        'upload_title': "📁 Data Upload (Click to expand/collapse)",
        'upload_help': "Upload Pick report, MARM report, TO details, and optional Manual Override.",
        'info_users': "💡 Excluded **{} system lines** (UIDJ5089, UIH25501).",
        'info_clean': "💡 1 move counted for **{} lines** of 'X' (Applies ONLY to PI_PL_FU, PI_PL_FUOE).",
        'info_manual': "✅ Loaded manual packaging for **{} unique materials**.",
        'sidebar_title': "⚙️ Algorithm Config",
        'weight_label': "Weight limit for 1-by-1 pick (kg)",
        'dim_label': "Dimension limit for 1-by-1 (cm)",
        'hmat_label': "Max pieces per grab",
        'exclude_label': "Exclude materials:",
        'sec_ratio': "🎯 Data Reliability & Source",
        'ratio_desc': "Data foundation (SAP Data Quality indicator):",
        'ratio_moves': "Share of total MOVEMENTS:",
        'ratio_exact': "Exact (Known packaging)",
        'ratio_miss': "Estimates (Missing packaging)",
        'logic_explain_title': "ℹ️ How does the app calculate physical moves?",
        'logic_explain_text': """The algorithm simulates real-life picker effort:
        1. **Full Units (X):** In Full Pallet queues (FU/FUOE), a 'X' removal counts as 1 move (forklift handling).
        2. **Boxes (Exact):** In other queues, the algorithm looks for full boxes based on Master data. Each box = 1 move.
        3. **Loose Pieces (Exact):** If packaging is known, remaining pieces are divided by the 'per grab' factor. Heavy or large items are picked 1-by-1.
        4. **Estimates (Miss):** If SAP lacks packaging data, a safety workload estimate is applied to prevent underestimating effort.""",
        'exp_missing_data': "🔍 Show materials with missing box data (Estimates Leaderboard)",
        'sec_queue_title': "📊 Average Workload by Queue",
        'filter_month': "📅 Filter by month:",
        'all_months': "All months",
        'all_queues': "All Queues combined",
        'unknown': "Unknown",
        'q_col_queue': "Queue",
        'q_col_desc': "Description",
        'q_col_to': "TOs",
        'q_col_orders': "Orders",
        'q_col_loc': "Avg Locs",
        'q_col_pcs': "Avg Pieces",
        'q_col_mov_loc': "Moves per Loc",
        'q_col_exact_loc': "Exact per Loc",
        'q_pct_exact': "% Exact",
        'q_col_miss_loc': "Estimate per Loc",
        'q_pct_miss': "% Estimate",
        'sec_queue_top_title': "🏆 TOP 100 Materials by Queue",
        'q_select': "Show TOP 100 for:",
        'sec1_title': "🎯 Pallet Order Analysis (Mix Pallet)",
        'pallets_clean_info': "*(Calculated only from PI_PL and PI_PL_OE)*",
        'm_orders': "Orders",
        'm_qty': "Avg Pcs / Order",
        'm_pos': "Avg Bins / Order",
        'm_mov_loc': "Avg Physical Moves per Loc",
        'exp_detail_title': "Show Orders Table",
        'col_mat': "Material",
        'col_qty': "Total Pieces",
        'col_mov': "Moves",
        'col_mov_exact': "Moves (Exact)",
        'col_mov_miss': "Moves (Estimates)",
        'col_wgt': "Weight (kg)",
        'col_max_dim': "Max Dim (cm)",
        'audit_title': "🎲 Detailed Audit Report (Random samples)",
        'audit_phys_moves': "Physical moves",
        'audit_gen_btn': "Generate Audit Report (5 TOs per Queue)",
        'sec3_title': "🔍 Master Data Viewer",
        'search_label': "Check specific material data:",
        'tab_dashboard': "📊 Dashboard & Queue",
        'tab_pallets': "📦 Pallet Orders",
        'tab_top': "🏆 TOP Materials",
        'tab_audit': "🔍 Tools & Audit",
        'col_lines': "Lines",
        'btn_download': "📥 Download Comprehensive Report (Excel)",
        'no_orders': "No orders found.",
        'err_pick': "Error: Pick report not found in uploads."
    }
}

# --- POMOCNÉ FUNKCE ---
def t(key):
    return TEXTS[st.session_state.lang][key]

def get_match_key(val):
    v = str(val).strip().upper()
    if '.' in v and v.replace('.', '').isdigit():
        return v.rstrip('0').rstrip('.')
    return v

# --- OPTIMALIZOVANÝ VÝPOČET ---
def fast_compute_moves(qty_s, queue_s, su_s, box_s, w_s, d_s, v_lim, d_lim, h_lim):
    """Vektorizovaná logika výpočtu pohybů pro maximální rychlost"""
    r_total, r_exact, r_miss = [], [], []
    for qty, q, su, boxes, w, d in zip(qty_s, queue_s, su_s, box_s, w_s, d_s):
        if qty <= 0:
            r_total.append(0); r_exact.append(0); r_miss.append(0)
            continue
            
        # Pravidlo X (Plné jednotky)
        if str(q).upper() in ('PI_PL_FU', 'PI_PL_FUOE') and str(su).strip().upper() == 'X':
            r_total.append(1); r_exact.append(1); r_miss.append(0)
            continue
            
        pb = pok = pmiss = 0
        zbytek = qty
        
        # 1. Krabice (vždy Exact)
        if boxes:
            for b in boxes:
                if b > 1 and zbytek >= b:
                    pb += int(zbytek // b)
                    zbytek %= b
                    
        # 2. Volné kusy
        if zbytek > 0:
            # Pokud je těžké nebo velké -> po 1 ks
            if w >= v_lim or d >= d_lim:
                p = zbytek
            else:
                p = int((zbytek + h_lim - 1) // h_lim)
                
            if not boxes: # Chybí Master data -> Odhad
                pmiss += p
            else:
                pok += p
                
        r_total.append(pb + pok + pmiss)
        r_exact.append(pb + pok)
        r_miss.append(pmiss)
        
    return r_total, r_exact, r_miss

# ==========================================
# 3. HLAVNÍ LOGIKA APLIKACE
# ==========================================
def main():
    # Hlavička
    col_t, col_l = st.columns([8, 1])
    with col_t:
        st.markdown(f"<div class='main-header'>{t('title')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='sub-header'>{t('desc')}</div>", unsafe_allow_html=True)
    with col_l:
        if st.button(t('switch_lang')):
            st.session_state.lang = 'en' if st.session_state.lang == 'cs' else 'cs'
            st.rerun()

    # Sidebar konfigurace
    st.sidebar.header(t('sidebar_title'))
    limit_vahy = st.sidebar.number_input(t('weight_label'), 0.1, 20.0, 2.0, 0.5)
    limit_rozmeru = st.sidebar.number_input(t('dim_label'), 1.0, 200.0, 15.0, 1.0)
    kusy_na_hmat = st.sidebar.slider(t('hmat_label'), 1, 20, 1) # Default 1 dle požadavku
    
    with st.expander(t('upload_title'), expanded=True):
        uploaded_files = st.file_uploader(t('upload_help'), type=['csv', 'xlsx'], accept_multiple_files=True)

    if uploaded_files:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        df_pick, df_marm, df_manual, df_queue = None, None, None, None

        # --- NAČÍTÁNÍ ---
        status_text.markdown("**🔄 Načítání a čtení souborů...**")
        progress_bar.progress(10)
        
        for file in uploaded_files:
            fname = file.name.lower()
            temp_df = pd.read_csv(file, dtype=str) if fname.endswith('.csv') else pd.read_excel(file, dtype=str)
            
            if 'Delivery' in temp_df.columns and 'Act.qty (dest)' in temp_df.columns:
                df_pick = temp_df
            elif 'Numerator' in temp_df.columns and 'Alternative Unit of Measure' in temp_df.columns:
                df_marm = temp_df
            elif 'Queue' in temp_df.columns:
                df_queue = temp_df
            elif len(temp_df.columns) >= 2:
                df_manual = temp_df

        if df_pick is None:
            st.error(t('err_pick'))
            return

        # --- ČIŠTĚNÍ A PÁROVÁNÍ ---
        status_text.markdown("**⚙️ Čištění dat a párování Queue...**")
        progress_bar.progress(30)
        
        df_pick['Material'] = df_pick['Material'].astype(str).str.strip()
        df_pick['Match_Key'] = df_pick['Material'].apply(get_match_key)
        df_pick['Qty'] = pd.to_numeric(df_pick['Act.qty (dest)'], errors='coerce').fillna(0)
        df_pick['Source Storage Bin'] = df_pick.get('Source Storage Bin', df_pick.get('Storage Bin', 'N/A'))
        
        # Vyloučení adminů
        mask_admins = df_pick['User'].isin(['UIDJ5089', 'UIH25501']) if 'User' in df_pick.columns else pd.Series(False, index=df_pick.index)
        num_removed_admins = mask_admins.sum()
        df_pick = df_pick[~mask_admins].copy()

        # Párování Queue z TO Details
        if df_queue is not None:
            q_map = df_queue.drop_duplicates('Transfer Order Number').set_index('Transfer Order Number')['Queue'].to_dict()
            df_pick['Queue'] = df_pick['Transfer Order Number'].map(q_map).fillna('N/A')
            d_map = df_queue.drop_duplicates('Transfer Order Number').set_index('Transfer Order Number')['Confirmation Date'].to_dict()
            df_pick['Date'] = df_pick['Transfer Order Number'].map(d_map)
        else:
            df_pick['Queue'], df_pick['Date'] = 'N/A', np.nan
        
        df_pick['Month'] = pd.to_datetime(df_pick['Date'], errors='coerce').dt.to_period('M').astype(str).replace('NaT', t('unknown'))
        df_pick['Removal of total SU'] = df_pick['Removal of total SU'].fillna('').astype(str).str.strip().upper()

        # --- MASTER DATA ---
        status_text.markdown("**📦 Zpracování Master Dat a balení...**")
        progress_bar.progress(60)
        
        box_dict, weight_dict, dim_dict = {}, {}, {}
        if df_marm is not None:
            df_marm['Match_Key'] = df_marm['Material'].apply(get_match_key)
            df_marm['Numerator'] = pd.to_numeric(df_marm['Numerator'], errors='coerce').fillna(0)
            
            # Krabice z MARM
            box_dict = df_marm[df_marm['Alternative Unit of Measure'].isin(['AEK','KAR','KART','PAK','VPE','CAR','BLO'])].groupby('Match_Key')['Numerator'].apply(lambda x: sorted([int(float(i)) for i in x if float(i)>1], reverse=True)).to_dict()
            
            # Váha a rozměr
            st_data = df_marm[df_marm['Alternative Unit of Measure'].isin(['ST','PCE','KS'])]
            weight_dict = st_data.groupby('Match_Key')['Gross Weight'].first().astype(float).to_dict()
            dim_dict = st_data.set_index('Match_Key')[['Length','Width','Height']].astype(float).max(axis=1).to_dict()

        # Ruční ověření
        if df_manual is not None:
            for _, r in df_manual.iterrows():
                m_key = get_match_key(r[0])
                nums = [int(x) for x in re.findall(r'\d+', str(r[1])) if int(x)>1]
                if nums: box_dict[m_key] = sorted(nums, reverse=True)

        # --- VÝPOČET ---
        status_text.markdown("**🤖 Simulace fyzické zátěže...**")
        progress_bar.progress(85)
        
        df_pick['Box_Sizes_List'] = df_pick['Match_Key'].map(box_dict)
        df_pick['W'] = df_pick['Match_Key'].map(weight_dict).fillna(0)
        df_pick['D'] = df_pick['Match_Key'].map(dim_dict).fillna(0)

        t_total, t_exact, t_miss = fast_compute_moves(
            df_pick['Qty'].values, df_pick['Queue'].values, df_pick['Removal of total SU'].values,
            df_pick['Box_Sizes_List'].values, df_pick['W'].values, df_pick['D'].values,
            limit_vahy, limit_rozmeru, kusy_na_hmat
        )
        
        df_pick['Pohyby_Rukou'] = t_total
        df_pick['Pohyby_Exact'] = t_exact
        df_pick['Pohyby_Loose_Miss'] = t_miss
        df_pick['Celkova_Vaha_KG'] = df_pick['Qty'] * df_pick['W']

        progress_bar.progress(100)
        time.sleep(0.3)
        progress_bar.empty()
        status_text.empty()

        # Banner Info
        c_i1, c_i2 = st.columns(2)
        if num_removed_admins > 0: c_i1.info(t('info_users').format(num_removed_admins))
        x_c = ((df_pick['Removal of total SU'] == 'X') & (df_pick['Queue'].str.contains('FU', na=False))).sum()
        if x_c > 0: c_i2.warning(t('info_clean').format(x_c))

        # ==========================================
        # 4. ZOBRAZENÍ TABŮ
        # ==========================================
        tab_dash, tab_pallets, tab_top, tab_audit = st.tabs([t('tab_dashboard'), t('tab_pallets'), t('tab_top'), t('tab_audit')])

        # --- TAB 1: DASHBOARD ---
        with tab_dash:
            tot_m = df_pick['Pohyby_Rukou'].sum()
            if tot_m > 0:
                st.subheader(t('sec_ratio'))
                st.write(t('ratio_desc'))
                st.markdown(f"**{t('ratio_moves')}**")
                
                c_r1, c_r2 = st.columns(2)
                c_r1.metric(t('ratio_exact'), f"{(df_pick['Pohyby_Exact'].sum()/tot_m*100):.1f} %", f"{df_pick['Pohyby_Exact'].sum():,.0f}")
                c_r2.metric(t('ratio_miss'), f"{(df_pick['Pohyby_Loose_Miss'].sum()/tot_m*100):.1f} %", f"{df_pick['Pohyby_Loose_Miss'].sum():,.0f}", delta_color="inverse")
                
                with st.expander(t('logic_explain_title')):
                    st.info(t('logic_explain_text'))

            st.divider()
            st.subheader(t('sec_queue_title'))
            
            sel_m = st.selectbox(t('filter_month'), [t('all_months')] + sorted(df_pick['Month'].unique().tolist()))
            df_f = df_pick if sel_m == t('all_months') else df_pick[df_pick['Month'] == sel_m]
            
            # Agregace Queue
            q_agg = df_f.groupby('Queue').agg(
                to_count=('Transfer Order Number','nunique'),
                orders=('Delivery','nunique'),
                loc_sum=('Source Storage Bin','nunique'),
                mov_sum=('Pohyby_Rukou','sum'),
                exact_sum=('Pohyby_Exact','sum'),
                miss_sum=('Pohyby_Loose_Miss','sum'),
                pcs_sum=('Qty','sum')
            ).reset_index()
            
            q_agg['Popis'] = q_agg['Queue'].map(QUEUE_DESC).fillna('')
            q_agg[t('q_col_loc')] = q_agg['loc_sum'] / q_agg['to_count']
            q_agg[t('q_col_pcs')] = q_agg['pcs_sum'] / q_agg['to_count']
            q_agg[t('q_col_mov_loc')] = q_agg['mov_sum'] / q_agg['loc_sum']
            q_agg[t('q_col_exact_loc')] = q_agg['exact_sum'] / q_agg['loc_sum']
            q_agg[t('q_col_miss_loc')] = q_agg['miss_sum'] / q_agg['loc_sum']
            q_agg[t('q_pct_exact')] = (q_agg['exact_sum'] / q_agg['mov_sum'] * 100)
            q_agg[t('q_pct_miss')] = (q_agg['miss_sum'] / q_agg['mov_sum'] * 100)

            # Tabulka Queue
            display_cols = ['Queue', 'Popis', 'to_count', 'orders', t('q_col_loc'), t('q_col_pcs'), 
                           t('q_col_mov_loc'), t('q_col_exact_loc'), t('q_pct_exact'), t('q_col_miss_loc'), t('q_pct_miss')]
            
            st.dataframe(
                q_agg[display_cols].style.format({c: "{:.1f}" for c in display_cols if 'Prům' in c or 'Avg' in c or 'na lokaci' in c or 'per Loc' in c} | {c: "{:.1f} %" for c in display_cols if '%' in c}),
                use_container_width=True, hide_index=True
            )
            
            st.markdown(f"**{t('q_col_mov_loc')} (Chart)**")
            st.bar_chart(q_agg.set_index('Queue')[t('q_col_mov_loc')])

        # --- TAB 2: PALETOVÉ ZAKÁZKY ---
        with tab_pallets:
            st.subheader(t('sec1_title'))
            st.caption(t('pallets_clean_info'))
            
            # Striktní filtrace Queue dle požadavku
            allowed_q = ['PI_PL (Mix)', 'PI_PL (Total)', 'PI_PL (Single)', 'PI_PL_OE (Mix)', 'PI_PL_OE (Total)', 'PI_PL_OE (Single)']
            df_pal = df_pick[df_pick['Queue'].isin(allowed_q)]
            
            if not df_pal.empty:
                pal_agg = df_pal.groupby('Delivery').agg(
                    mat_count=('Material','nunique'),
                    mov=('Pohyby_Rukou','sum'),
                    ex=('Pohyby_Exact','sum'),
                    mi=('Pohyby_Loose_Miss','sum'),
                    locs=('Source Storage Bin','nunique'),
                    qty=('Qty','sum')
                )
                # Pouze 1-materiálové zakázky
                pal_f = pal_agg[pal_agg['mat_count'] == 1].copy()
                
                if not pal_f.empty:
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric(t('m_orders'), f"{len(pal_f):,}")
                    c2.metric(t('m_qty'), f"{pal_f['qty'].mean():.1f}")
                    c3.metric(t('m_pos'), f"{pal_f['locs'].mean():.2f}")
                    c4.metric(t('m_mov_loc'), f"{(pal_f['mov'].sum()/pal_f['locs'].sum()):.1f}")
                    
                    st.divider()
                    st.write(t('exp_detail_title'))
                    
                    # Detailní tabulka
                    pal_det = df_pal[df_pal['Delivery'].isin(pal_f.index)].groupby('Material').agg(
                        cnt=('Delivery','nunique'),
                        qty=('Qty','sum'),
                        mov=('Pohyby_Rukou','sum'),
                        ex=('Pohyby_Exact','sum'),
                        mi=('Pohyby_Loose_Miss','sum'),
                        weight=('Celkova_Vaha_KG','sum'),
                        dim=('D', 'max')
                    ).sort_values('mov', ascending=False).reset_index()
                    
                    pal_det.columns = [t('col_mat'), t('m_orders'), t('col_qty'), t('col_mov'), t('col_mov_exact'), t('col_mov_miss'), t('col_wgt'), t('col_max_dim')]
                    st.dataframe(pal_det, use_container_width=True, hide_index=True)
                else:
                    st.warning(t('no_orders'))
            else:
                st.warning(t('no_orders'))

        # --- TAB 3: TOP MATERIÁLY ---
        with tab_top:
            st.subheader(t('sec_queue_top_title'))
            q_list = [t('all_queues')] + sorted(df_pick['Queue'].unique().tolist())
            sel_q = st.selectbox(t('q_select'), q_list)
            
            df_t = df_pick if sel_q == t('all_queues') else df_pick[df_pick['Queue'] == sel_q]
            
            top_res = df_t.groupby('Material').agg(
                lines=('Material','count'),
                qty=('Qty','sum'),
                mov=('Pohyby_Rukou','sum'),
                ex=('Pohyby_Exact','sum'),
                mi=('Pohyby_Loose_Miss','sum')
            ).sort_values('mov', ascending=False).head(100).reset_index()
            
            top_res.columns = [t('col_mat'), t('col_lines'), t('col_qty'), t('col_mov'), t('col_mov_exact'), t('col_mov_miss')]
            
            col_g1, col_g2 = st.columns([2, 1])
            with col_g1:
                st.dataframe(top_res, use_container_width=True, hide_index=True)
            with col_g2:
                st.bar_chart(top_res.set_index(t('col_mat'))[t('col_mov')])

            st.divider()
            st.subheader(t('exp_missing_data'))
            all_miss = df_pick.groupby('Material').agg(miss=('Pohyby_Loose_Miss','sum'), lines=('Material','count')).reset_index()
            miss_only = all_miss[all_miss['miss'] > 0].sort_values('miss', ascending=False).head(50)
            st.dataframe(miss_only, use_container_width=True)

        # --- TAB 4: AUDIT & NÁSTROJE ---
        with tab_audit:
            col_au1, col_au2 = st.columns([3, 2])
            
            with col_au1:
                st.subheader(t('audit_title'))
                if st.button(t('audit_gen_btn'), type="primary"):
                    valid_qs = sorted([q for q in df_pick['Queue'].unique() if q != 'N/A'])
                    for q in valid_qs:
                        with st.expander(f"📁 Queue: {q}", expanded=False):
                            q_data = df_pick[df_pick['Queue'] == q]
                            samples = q_data.sample(min(5, len(q_data)))
                            
                            for _, r in samples.iterrows():
                                st.markdown(f"**TO: `{r['Transfer Order Number']}`** | Mat: `{r['Material']}` | Qty: `{r['Qty']} pcs`")
                                st.caption(f"Bin: `{r['Source Storage Bin']}` | Delivery: `{r['Delivery']}`")
                                
                                # Logický rozpad pro audit
                                q_str = str(r['Queue']).upper()
                                su_str = str(r['Removal of total SU']).upper()
                                rem = r['Qty']
                                
                                if su_str == 'X' and 'FU' in q_str:
                                    st.info(f"➡️ SU 'X' detected in {q_str}. Movement count: **1**")
                                else:
                                    if su_str == 'X':
                                        st.caption(f"*(Ignored 'X' marker because queue {q_str} is not Full Pallet)*")
                                    
                                    if r['Box_Sizes_List']:
                                        for b in r['Box_Sizes_List']:
                                            if b > 1 and rem >= b:
                                                st.write(f"• {int(rem // b)}x Full Box ({b} pcs) = {int(rem // b)} moves")
                                                rem %= b
                                    
                                    if rem > 0:
                                        if r['W'] >= limit_vahy or r['D'] >= limit_rozmeru:
                                            st.warning(f"• {int(rem)}x Heavy/Large piece ({r['W']:.2f}kg, {r['D']:.0f}cm) = {int(rem)} moves")
                                        else:
                                            grabs = int((rem + kusy_na_hmat - 1) // kusy_na_hmat)
                                            st.success(f"• {int(rem)}x Loose piece (Grab factor {kusy_na_hmat}) = {grabs} moves")
                                
                                st.markdown(f"> **{t('audit_phys_moves')}: `{r['Pohyby_Rukou']}`** (Exact: {r['Pohyby_Exact']}, Miss: {r['Pohyby_Loose_Miss']})")
                                st.write("---")

            with col_au2:
                st.subheader(t('sec3_title'))
                mat_s = st.selectbox(t('search_label'), [""] + sorted(df_pick['Material'].unique().tolist()))
                if mat_s:
                    mk = get_match_key(mat_s)
                    st.info(f"**{mat_s}**")
                    st.write(f"Weight: **{weight_dict.get(mk, 0):.3f} kg**")
                    st.write(f"Max Dim: **{dim_dict.get(mk, 0):.1f} cm**")
                    st.write(f"Packaging (Master): `{box_dict.get(mk, 'N/A')}`")
                    
                    if df_marm is not None:
                        st.write("Raw MARM data:")
                        st.dataframe(df_marm[df_marm['Match_Key'] == mk], hide_index=True)

            # EXPORT EXCEL
            st.divider()
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # List 1: Metodika
                pd.DataFrame({
                    "Parameter": ["Weight Limit", "Dim Limit", "Pieces per Grab", "Admin Excluded"],
                    "Value": [f"{limit_vahy} kg", f"{limit_rozmeru} cm", f"{kusy_na_hmat} ks", num_removed_admins]
                }).to_excel(writer, index=False, sheet_name='Methodology')
                
                # List 2: Queue
                q_agg.to_excel(writer, index=False, sheet_name='Queue_Summary')
                
                # List 3: TOP 100
                top_res.to_excel(writer, index=False, sheet_name='Top_100_Materials')
                
            st.download_button(
                label=t('btn_download'),
                data=buffer.getvalue(),
                file_name=f"Picking_Analysis_{time.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

if __name__ == "__main__":
    main()
