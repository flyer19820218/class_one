import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from matplotlib import font_manager

st.set_page_config(page_title="班級成績查詢系統", layout="centered", page_icon="🎓")

# ==========================================
# 區塊 1: 本地字體初始化 (極簡版)
# ==========================================
FONT_NAME = "NotoSansTC-Regular.ttf"

if os.path.exists(FONT_NAME):
    font_manager.fontManager.addfont(FONT_NAME)
    dynamic_font_name = font_manager.FontProperties(fname=FONT_NAME).get_name()
    plt.rcParams['font.sans-serif'] = [dynamic_font_name, 'Microsoft JhengHei', 'Arial']
    plt.rcParams['axes.unicode_minus'] = False
else:
    st.warning("⚠️ 找不到字體檔！雷達圖的中文可能會變成方塊。請確認 NotoSansTC-Regular.ttf 與 app.py 放在同一個資料夾。")

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
    if not csv_url:
        return None
    df = pd.read_csv(csv_url)
    df.columns = df.columns.str.strip()
    return df

# ==========================================
# 區塊 2: 網頁雷達圖產生器
# ==========================================
def draw_web_radar_chart(labels, pr_scores):
    # 防呆：確保所有 PR 分數都是有效的數字，避免雷達圖變形
    pr_scores = [0 if pd.isna(x) else x for x in pr_scores]
    
    num_vars = len(labels)
    if num_vars == 0:
        return plt.figure()
        
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
# 區塊 3: 網頁主介面
# ==========================================
st.title("🏫 班級成績查詢系統")

col_mode, col_grade = st.columns(2)
with col_mode:
    mode = st.radio("⚙️ 系統模式：", ["單次段考查詢", "總目錄主控表"])
with col_grade:
    # 【聽您的，把年級選擇按鈕加回來了！】
    grade_option = st.radio(
        "📊 選擇學生年級：",
        ["一年級 (五科版：國英數自社)", "二/三年級 (七科版：含史地公)"],
        index=1
    )

has_7_subjects = (grade_option == "二/三年級 (七科版：含史地公)")

target_sheet_url = None

if mode == "總目錄主控表":
    st.info("💡 請確保您的總目錄包含「考試名稱」與「該次成績單的 Google 試算表網址」兩個欄位。")
    MASTER_MENU_URL = st.text_input("🔗 請輸入「總目錄主控表」網址：")
    if MASTER_MENU_URL:
        try:
            df_menu = load_data(MASTER_MENU_URL)
            if df_menu is None:
                st.error("❌ 無法讀取總目錄，請確認網址與權限。")
            elif '考試名稱' not in df_menu.columns:
                st.error("❌ 找不到「考試名稱」欄位！請確認您貼上的是總目錄。")
            else:
                exam_options = df_menu['考試名稱'].dropna().tolist()
                selected_exam = st.selectbox("📅 請選擇考試項目：", exam_options)
                target_sheet_url = df_menu[df_menu['考試名稱'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
        except Exception as e:
            st.error(f"❌ 總目錄讀取失敗：{e}")
else:
    target_sheet_url = st.text_input("🔗 請輸入「單次段考成績單」網址：", placeholder="https://docs.google.com/spreadsheets/d/...")

if target_sheet_url:
    try:
        raw_df = load_data(target_sheet_url)
        if raw_df is not None:
            raw_df['_is_student'] = pd.to_numeric(raw_df['座號'], errors='coerce').notna()
            df_students = raw_df[raw_df['_is_student']].copy()
            df_students['座號'] = df_students['座號'].astype(int)
            df_students = df_students[df_students['座號'] > 0]
            
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
            
            # 完全依照上方老師選擇的年級，不再亂猜測
            if has_7_subjects:
                display_subs = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                radar_subs = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
            else:
                display_subs = ['國文', '英文', '數學', '自然', '社會']
                radar_subs = ['國文', '英文', '數學', '自然', '社會']
            
            # 強制補齊所有必要欄位（避免試算表缺漏導致閃退）
            for col in cols_to_extract:
                if col not in df_students.columns:
                    df_students[col] = 0
                    
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
                df_students[f'{sub}_PR'] = df_students[sub].rank(pct=True) * 100
            
            df_students = df_students.sort_values(by='座號').reset_index(drop=True)
            
            student_seats = sorted(df_students['座號'].tolist())
            selected_seat = st.selectbox("🔢 請選擇您的座號：", student_seats)
            
            if selected_seat:
                st.markdown("---")
                updated_student_df = df_students[df_students['座號'] == selected_seat].iloc[0]
                
                st.markdown(f"### 👤 座號 {selected_seat} 號 【{updated_student_df['姓名']}】 的成績報告")
                
                col1, col2 = st.columns([1.1, 1.0])
                
                with col1:
                    st.write("**📊 各科成績明細：**")
                    for s in display_subs:
                        if s in updated_student_df:
                            score = float(updated_student_df[s]) if pd.notna(updated_student_df[s]) else 0.0
                            stat = stats_dict.get(s, {})
                            avg = float(stat.get('平均', 0)) if pd.notna(stat.get('平均')) else 0.0
                            high = float(stat.get('高標', 0)) if pd.notna(stat.get('高標')) else 0.0
                            
                            if s in ['歷史', '地理', '公民']:
                                st.text(f"  └ {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                            else:
                                st.text(f"  {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                    
                    st.markdown("---")
                    total_score = float(updated_student_df['總分']) if pd.notna(updated_student_df['總分']) else 0.0
                    rank = int(float(updated_student_df['名次'])) if pd.notna(updated_student_df['名次']) else 0
                    total_pr = float(updated_student_df['總PR']) if pd.notna(updated_student_df['總PR']) else 0.0
                    
                    st.metric(label="五科總分", value=f"{total_score:.2f}")
                    st.write(f"🏆 **班級名次**：第 {rank} 名")
                    st.write(f"📈 **總分 PR 值**：{total_pr:.2f}")
                    
                with col2:
                    # 依據年級選項，繪製完美的多邊形雷達圖
                    pr_scores = [updated_student_df[f'{s}_PR'] for s in radar_subs]
                    fig = draw_web_radar_chart(radar_subs, pr_scores)
                    st.pyplot(fig)
                        
    except Exception as e:
        st.error(f"❌ 系統發生預期外的錯誤，請確認欄位是否正確。錯誤資訊: {e}")
