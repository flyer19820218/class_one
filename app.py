import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager

st.set_page_config(page_title="801 專屬成績大數據主控台", layout="centered", page_icon="🎓")

# ==========================================
# 區塊 1: 字體與環境初始化
# ==========================================
FONT_NAME = "NotoSansTC-Regular.ttf"

if os.path.exists(FONT_NAME):
    font_manager.fontManager.addfont(FONT_NAME)
    dynamic_font_name = font_manager.FontProperties(fname=FONT_NAME).get_name()
    plt.rcParams['font.sans-serif'] = [dynamic_font_name, 'Microsoft JhengHei', 'Arial']
    plt.rcParams['axes.unicode_minus'] = False
else:
    st.warning("⚠️ 找不到字體檔！圖表的中文可能會變成方塊。請確認 NotoSansTC-Regular.ttf 與 app.py 放在同一個資料夾。")

def get_google_sheet_csv_url(url):
    try:
        if "export?format=csv" in url: return url
        import re
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            gid_match = re.search(r'gid=([0-9]+)', url)
            gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
            return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv{gid_param}"
    except: pass
    return None

@st.cache_data(ttl=60)
def load_data(url):
    csv_url = get_google_sheet_csv_url(url)
    if not csv_url: return None
    try:
        df = pd.read_csv(csv_url)
        df.columns = df.columns.str.strip()
        return df
    except: return None

# ==========================================
# 區塊 2: 歷史跨考試數據撈取引擎 (核心功能)
# ==========================================
def fetch_student_history(df_menu, seat_num):
    history_records = []
    
    for _, row in df_menu.iterrows():
        exam_name = str(row['考試名稱']).strip()
        file_url = str(row['該次成績單的 Google 試算表網址']).strip()
        
        df_exam = load_data(file_url)
        if df_exam is not None:
            df_exam['_is_student'] = pd.to_numeric(df_exam['座號'], errors='coerce').notna()
            df_stud = df_exam[df_exam['_is_student']].copy()
            df_stud['座號'] = df_stud['座號'].astype(int)
            
            # 找到該座號學生
            match = df_stud[df_stud['座號'] == seat_num]
            if not match.empty:
                s_row = match.iloc[0]
                
                # 基礎四科數值轉換
                g_num = lambda col: pd.to_numeric(s_row.get(col, 0), errors='coerce') if pd.notna(s_row.get(col, 0)) else 0
                
                chinese = g_num('國文')
                english = g_num('英文')
                math = g_num('數學')
                science = g_num('自然')
                
                # 解決一二年級銜接核心：如果有一年級的「社會」，直接用；若有「史地公」，取平均融合成社會分數
                if '歷史' in df_exam.columns and '地理' in df_exam.columns and '公民' in df_exam.columns:
                    social = (g_num('歷史') + g_num('地理') + g_num('公民')) / 3
                else:
                    social = g_num('社會')
                    
                rank = g_num('名次')
                total_score = g_num('總分')
                
                history_records.append({
                    "考試名稱": exam_name,
                    "國文": chinese, "英文": english, "數學": math, "自然": science, "社會(含領域融合)": social,
                    "總分": total_score, "名次": rank
                })
                
    return pd.DataFrame(history_records)

# ==========================================
# 區塊 3: 網頁主介面 (已直接綁定老師的總表)
# ==========================================
st.title("🎓 801 班級專屬成績查詢系統")
st.markdown("已鎖定專屬成績主控總表。學生可查詢單次成績，導師可檢視多期趨勢分析。")

# 🔒 【已直接寫進您的總表網址，一勞永逸】
MASTER_MENU_URL = "https://docs.google.com/spreadsheets/d/1FL-orK8H_oDrLuDg1pAijJICjHzPxJLcPIEngF0wko8/edit?gid=0#gid=0"

df_menu = load_data(MASTER_MENU_URL)

if df_menu is not None:
    df_menu.columns = df_menu.columns.str.strip()
    exam_options = df_menu['考試名稱'].dropna().tolist()
    
    # 建立上方大區塊切換
    tab1, tab2 = st.tabs(["🔍 學生個人單次查詢", "📈 導師動態趨勢分析後台"])
    
    with tab1:
        selected_exam = st.selectbox("📅 請選擇要查看的考試項目：", exam_options, key="student_exam")
        target_sheet_url = df_menu[df_menu['考試名稱'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
        
        raw_df = load_data(target_sheet_url)
        if raw_df is not None:
            raw_df['_is_student'] = pd.to_numeric(raw_df['座號'], errors='coerce').notna()
            df_students = raw_df[raw_df['_is_student']].copy()
            df_students['座號'] = df_students['座號'].astype(int)
            
            student_seats = sorted(df_students[df_students['座號'] > 0]['座號'].tolist())
            selected_seat = st.selectbox("🔢 請選擇您的座號：", student_seats, key="student_seat")
            
            if selected_seat:
                stud_data = df_students[df_students['座號'] == selected_seat].iloc[0]
                st.markdown(f"### 👤 座號 {selected_seat} 號 【{stud_data['姓名']}】 的成績報告")
                
                # 簡單列出該次成績
                cols = st.columns(4)
                fields = ['國文', '英文', '數學', '自然']
                for idx, f in enumerate(fields):
                    if f in stud_data:
                        cols[idx % 4].metric(f, f"{pd.to_numeric(stud_data[f], errors='coerce'):.1f}")
                
                st.markdown("---")
                c_rank, c_total = st.columns(2)
                with c_rank: st.metric("班級名次", f"第 {int(pd.to_numeric(stud_data.get('名次', 0), errors='coerce'))} 名")
                with c_total: st.metric("五科總分", f"{pd.to_numeric(stud_data.get('總分', 0), errors='coerce'):.1f}")

    with tab2:
        st.markdown("### 📊 學生歷史跨學期趨勢追蹤")
        analysis_seat = st.selectbox("🔢 請選擇要分析追蹤的學生座號：", student_seats, key="analysis_seat")
        
        if analysis_seat:
            with st.spinner("正在跨檔案分析歷史大數據..."):
                df_history = fetch_student_history(df_menu, analysis_seat)
                
            if not df_history.empty:
                stud_name = df_students[df_students['座號'] == analysis_seat]['姓名'].values[0]
                st.markdown(f"#### 📈 {analysis_seat} 號 【{stud_name}】 學習軌跡綜合分析")
                
                # 功能 (1)：各科成績總平均和趨勢圖
                fig_score, ax_score = plt.subplots(figsize=(8, 4.5))
                sub_cols = ["國文", "英文", "數學", "自然", "社會(含領域融合)"]
                
                for sub in sub_cols:
                    ax_score.plot(df_history["考試名稱"], df_history[sub], marker='o', linewidth=2, label=sub)
                
                ax_score.set_title("學科成績走勢圖 (一二年級完美串聯版)", fontsize=14, fontweight='bold')
                ax_score.set_ylabel("分數", fontsize=12)
                ax_score.set_ylim(0, 105)
                ax_score.grid(True, linestyle=':', alpha=0.6)
                ax_score.legend(loc='lower left', ncol=5, fontsize=9)
                plt.xticks(rotation=15)
                st.pyplot(fig_score)
                
                # 功能 (2)：名次趨勢圖
                fig_rank, ax_rank = plt.subplots(figsize=(8, 3.5))
                ax_rank.plot(df_history["考試名稱"], df_history["名次"], marker='s', color='#E67E22', linewidth=2.5, label="班級名次")
                
                ax_rank.set_title("個人班級名次晉升軌跡 (曲線往上代表進步)", fontsize=14, fontweight='bold')
                ax_rank.set_ylabel("名次", fontsize=12)
                ax_rank.grid(True, linestyle=':', alpha=0.6)
                
                # 【重要細節】：反轉 Y 軸，讓第 1 名在最上面！
                ax_rank.invert_yaxis() 
                
                # 動態設定名次刻度，符合真實班級人數
                max_rank = int(df_history["名次"].max()) + 3
                ax_rank.set_ylim(max_rank, 1)
                
                plt.xticks(rotation=15)
                st.pyplot(fig_rank)
                
                # 數據表格備查
                with st.expander("👀 檢視完整歷史數據清單"):
                    st.dataframe(df_history)
            else:
                st.warning("找不到該位同學的歷史紀錄。")
else:
    st.error("❌ 無法載入您的主控總表，請確認該 Google Sheets 權限已開啟為「知道連結的人均可檢視」。")
