import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager

st.set_page_config(page_title="801 專屬成績大數據主控台", layout="centered", page_icon="🎓")

# ==========================================
# 區塊 1: 本地字體初始化
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

def check_is_grade_2_or_3(exam_name):
    if any(keyword in str(exam_name) for keyword in ["二上", "二下", "三上", "三下"]):
        return True
    return False

# ==========================================
# 區塊 2: 雷達圖產生器 (單次段考專用)
# ==========================================
def draw_web_radar_chart(labels, pr_scores):
    pr_scores = [0 if pd.isna(x) else x for x in pr_scores]
    num_vars = len(labels)
    if num_vars == 0: return plt.figure()
        
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    plot_scores = list(pr_scores) + [pr_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    display_labels = [f"{label}\n(PR)" for label in labels]
    ax.set_thetagrids(np.degrees(angles[:-1]), display_labels, fontsize=10)
    
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["25", "50", "75", "100"], color="grey", size=8)
    
    ax.plot(angles, [50]*len(angles), color='#C0504D', linewidth=1.5, linestyle='--', label='PR 50 (中位數)')
    ax.plot(angles, plot_scores, color='#4F81BD', linewidth=2, linestyle='solid', label='個人表現')
    ax.fill(angles, plot_scores, color='#4F81BD', alpha=0.35)
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fontsize=9, ncol=2)
    
    return fig

# ==========================================
# 區塊 3: 歷史跨考試數據撈取引擎 (已改為全面 PR 化)
# ==========================================
def fetch_student_history(df_menu, seat_num):
    history_records = []
    
    for _, row in df_menu.iterrows():
        exam_name = str(row['考試檔案名稱或別名'] if '考試檔案名稱或別名' in df_menu.columns else row['考試名稱']).strip()
        file_url = str(row['該次成績單的 Google 試算表網址']).strip()
        
        df_exam = load_data(file_url)
        if df_exam is not None:
            df_exam['_is_student'] = pd.to_numeric(df_exam['座號'], errors='coerce').notna()
            df_stud = df_exam[df_exam['_is_student']].copy()
            df_stud['座號'] = df_stud['座號'].astype(int)
            
            is_grade_2_3 = check_is_grade_2_or_3(exam_name)
            
            # 清理與轉換全班數據型態
            for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
                if col not in df_stud.columns: df_stud[col] = 0
                df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce').fillna(0)
            
            # 處理跨年級社會領域融合
            if is_grade_2_3:
                df_stud['社會_融合'] = (df_stud['歷史'] + df_stud['地理'] + df_stud['公民']) / 3
            else:
                df_stud['社會_融合'] = df_stud['社會']
            
            # 全自動動態計算全班各學科在該次考試的 PR 值
            for sub in ['國文', '英文', '數學', '自然', '社會_融合']:
                df_stud[f'{sub}_PR'] = df_stud[sub].rank(pct=True) * 100
            
            # 自動推算總分與名次
            if '總分' not in df_stud.columns:
                df_stud['總分'] = df_stud['國文'] + df_stud['英文'] + df_stud['數學'] + df_stud['自然'] + df_stud['社會_融合']
            df_stud['總分'] = pd.to_numeric(df_stud['總分'], errors='coerce').fillna(0)
            
            if '名次' not in df_stud.columns:
                df_stud['名次'] = df_stud['總分'].rank(ascending=False, method='min')
            
            # 擷取目標學生的數據
            match = df_stud[df_stud['座號'] == seat_num]
            if not match.empty:
                s_row = match.iloc[0]
                history_records.append({
                    "考試名稱": exam_name,
                    "國文": s_row['國文_PR'],
                    "英文": s_row['英文_PR'],
                    "數學": s_row['數學_PR'],
                    "自然": s_row['自然_PR'],
                    "社會(含領域融合)": s_row['社會_融合_PR'],
                    "總分": pd.to_numeric(s_row.get('總分', 0)),
                    "名次": pd.to_numeric(s_row.get('名次', 0))
                })
                
    return pd.DataFrame(history_records)

# ==========================================
# 區塊 4: 網頁主介面
# ==========================================
st.title("🎓 801 專屬成績大數據主控台")

mode = st.radio("⚙️ 系統模式：", ["總目錄主控表", "單次段考查詢 (不建議，無法使用歷史趨勢)"])

# 固定鎖定您的個人總表網址
MASTER_MENU_URL = "https://docs.google.com/spreadsheets/d/1FL-orK8H_oDrLuDg1pAijJICjHzPxJLcPIEngF0wko8/edit?gid=0#gid=0"

df_menu = load_data(MASTER_MENU_URL)

if df_menu is not None:
    # 相容欄位名稱
    if '考試檔案名稱或別名' in df_menu.columns:
        exam_options = df_menu['考試檔案名稱或別名'].dropna().tolist()
    else:
        exam_options = df_menu['考試名稱'].dropna().tolist()
    
    tab1, tab2 = st.tabs(["🔍 學生個人單次查詢", "📈 導師動態趨勢分析後台"])
    
    with tab1:
        if mode == "總目錄主控表":
            selected_exam = st.selectbox("📅 請選擇要查看的考試項目：", exam_options, key="student_exam")
            if '考試檔案名稱或別名' in df_menu.columns:
                target_sheet_url = df_menu[df_menu['考試檔案名稱或別名'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
            else:
                target_sheet_url = df_menu[df_menu['考試名稱'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
            has_7_subjects = check_is_grade_2_or_3(selected_exam)
        else:
            target_sheet_url = st.text_input("🔗 請輸入「單次段考成績單」網址：", placeholder="https://docs.google.com/spreadsheets/d/...")
            has_7_subjects = True
            
        if target_sheet_url:
            raw_df = load_data(target_sheet_url)
            if raw_df is not None:
                raw_df['_is_student'] = pd.to_numeric(raw_df['座號'], errors='coerce').notna()
                df_students = raw_df[raw_df['_is_student']].copy()
                df_students['座號'] = df_students['座號'].astype(int)
                
                df_bottom = raw_df[~raw_df['_is_student']].copy()
                
                stats_dict = {}
                cols_to_extract = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                for _, row in df_bottom.iterrows():
                    name = str(row['座號']).strip() if pd.notna(row['座號']) else str(row['姓名']).strip()
                    if '高標' in name:
                        for c in cols_to_extract:
                            if c in row:
                                if c not in stats_dict: stats_dict[c] = {}
                                stats_dict[c]['高標'] = pd.to_numeric(row[c], errors='coerce')
                    if '平均' in name or '均標' in name:
                        for c in cols_to_extract:
                            if c in row:
                                if c not in stats_dict: stats_dict[c] = {}
                                stats_dict[c]['平均'] = pd.to_numeric(row[c], errors='coerce')
                
                if has_7_subjects:
                    display_subs = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                    radar_subs = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
                else:
                    display_subs = ['國文', '英文', '數學', '自然', '社會']
                    radar_subs = ['國文', '英文', '數學', '自然', '社會']
                
                for col in cols_to_extract:
                    if col not in df_students.columns: df_students[col] = 0
                
                if '總分' not in df_students.columns:
                    if has_7_subjects:
                        df_students['總分'] = pd.to_numeric(df_students['國文'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['英文'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['數學'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['自然'], errors='coerce').fillna(0) + \
                                              (pd.to_numeric(df_students['歷史'], errors='coerce').fillna(0) + \
                                               pd.to_numeric(df_students['地理'], errors='coerce').fillna(0) + \
                                               pd.to_numeric(df_students['公民'], errors='coerce').fillna(0)) / 3
                    else:
                        df_students['總分'] = pd.to_numeric(df_students['國文'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['英文'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['數學'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['自然'], errors='coerce').fillna(0) + \
                                              pd.to_numeric(df_students['社會'], errors='coerce').fillna(0)
                
                df_students['總分'] = pd.to_numeric(df_students['總分'], errors='coerce').fillna(0)
                df_students['總PR'] = df_students['總分'].rank(pct=True) * 100
                
                if '名次' in df_students.columns:
                    df_students['名次'] = pd.to_numeric(df_students['名次'], errors='coerce').fillna(0)
                else:
                    df_students['名次'] = df_students['總分'].rank(ascending=False, method='min')
                
                for sub in radar_subs:
                    df_students[sub] = pd.to_numeric(df_students[sub], errors='coerce').fillna(0)
                    df_students[f'{sub}_PR_current'] = df_students[sub].rank(pct=True) * 100
                
                df_students = df_students.sort_values(by='座號').reset_index(drop=True)
                student_seats = sorted(df_students[df_students['座號'] > 0]['座號'].tolist())
                selected_seat = st.selectbox("🔢 請選擇您的座號：", student_seats, key="student_seat")
                
                if selected_seat:
                    st.markdown("---")
                    stud_data = df_students[df_students['座號'] == selected_seat].iloc[0]
                    st.markdown(f"### 👤 座號 {selected_seat} 號 【{stud_data['姓名']}】 的成績報告")
                    
                    col1, col2 = st.columns([1.1, 1.0])
                    
                    with col1:
                        st.write("**📊 各科成績明細：**")
                        for s in display_subs:
                            if s in stud_data:
                                score = float(stud_data[s]) if pd.notna(stud_data[s]) else 0.0
                                stat = stats_dict.get(s, {})
                                avg = float(stat.get('平均', 0)) if pd.notna(stat.get('平均')) else 0.0
                                high = float(stat.get('高標', 0)) if pd.notna(stat.get('高標')) else 0.0
                                
                                if s in ['歷史', '地理', '公民']:
                                    st.text(f"  └ {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                                else:
                                    st.text(f"  {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                        
                        st.markdown("---")
                        total_score = float(stud_data['總分'])
                        rank = int(float(stud_data['名次']))
                        total_pr = float(stud_data['總PR'])
                        
                        st.metric(label="五科總分", value=f"{total_score:.2f}")
                        st.write(f"🏆 **班級名次**：第 {rank} 名")
                        st.write(f"📈 **總分 PR 值**：{total_pr:.2f}")
                        
                    with col2:
                        pr_scores = [stud_data[f'{s}_PR_current'] for s in radar_subs]
                        fig = draw_web_radar_chart(radar_subs, pr_scores)
                        st.pyplot(fig)

    with tab2:
        st.markdown("### 📊 學生歷史跨學期趨勢追蹤")
        if 'student_seats' in locals() and student_seats:
            analysis_seat = st.selectbox("🔢 請選擇要分析追蹤的學生座號：", student_seats, key="analysis_seat")
            
            if analysis_seat:
                with st.spinner("正在跨檔案分析歷史大數據..."):
                    df_history = fetch_student_history(df_menu, analysis_seat)
                    
                if not df_history.empty:
                    stud_name = df_students[df_students['座號'] == analysis_seat]['姓名'].values[0]
                    st.markdown(f"#### 📈 {analysis_seat} 號 【{stud_name}】 學習軌跡綜合分析")
                    
                    # 功能 (1)：各科 PR 值趨勢圖 (已全面進化為 PR)
                    fig_score, ax_score = plt.subplots(figsize=(8, 4.5))
                    sub_cols = ["國文", "英文", "數學", "自然", "社會(含領域融合)"]
                    
                    for sub in sub_cols:
                        ax_score.plot(df_history["考試名稱"], df_history[sub], marker='o', linewidth=2, label=f"{sub} PR")
                    
                    ax_score.set_title("各科 PR 值走勢圖 (跨學期真實能力追蹤)", fontsize=14, fontweight='bold')
                    ax_score.set_ylabel("PR 值 (越大表現越優異)", fontsize=12)
                    ax_score.set_ylim(0, 105)
                    ax_score.grid(True, linestyle=':', alpha=0.6)
                    ax_score.legend(loc='lower left', ncol=5, fontsize=9)
                    plt.xticks(rotation=15)
                    st.pyplot(fig_score)
                    
                    # 功能 (2)：名次趨勢圖 (曲線往上代表進步)
                    fig_rank, ax_rank = plt.subplots(figsize=(8, 3.5))
                    ax_rank.plot(df_history["考試名稱"], df_history["名次"], marker='s', color='#E67E22', linewidth=2.5, label="班級名次")
                    
                    ax_rank.set_title("個人班級名次晉升軌跡 (曲線往上代表進步)", fontsize=14, fontweight='bold')
                    ax_rank.set_ylabel("名次", fontsize=12)
                    ax_rank.grid(True, linestyle=':', alpha=0.6)
                    ax_rank.invert_yaxis() 
                    
                    max_rank = int(df_history["名次"].max()) + 3
                    ax_rank.set_ylim(max_rank, 1)
                    
                    plt.xticks(rotation=15)
                    st.pyplot(fig_rank)
                    
                    with st.expander("👀 檢視完整歷史數據清單"):
                        st.dataframe(df_history)
                else:
                    st.warning("找不到該位同學的歷史紀錄。")
        else:
            st.warning("請先在第一頁載入正確的成績資訊。")
else:
    st.error("❌ 無法載入您的主控總表，請確認網址權限設定。")
