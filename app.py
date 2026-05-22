import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import requests
import os
import io
from matplotlib import font_manager

st.set_page_config(page_title="學生成績查詢系統", layout="centered", page_icon="🎓")

# ==========================================
# 區塊 1: 字體與環境初始化
# ==========================================
FONT_NAME = "NotoSansTC-Regular.ttf"
FONT_URL = "https://cdn.jsdelivr.net/gh/themoeway/noto-sans-tc-ttf@master/ttf/NotoSansTC-Regular.ttf"

@st.cache_resource
def init_fonts():
    if not os.path.exists(FONT_NAME):
        try:
            with st.spinner("首次啟動，正在載入雲端環境..."):
                response = requests.get(FONT_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
                with open(FONT_NAME, "wb") as f:
                    f.write(response.content)
        except:
            return False
    try:
        font_manager.fontManager.addfont(FONT_NAME)
        dynamic_font_name = font_manager.FontProperties(fname=FONT_NAME).get_name()
        plt.rcParams['font.sans-serif'] = [dynamic_font_name, 'Microsoft JhengHei', 'Arial']
        plt.rcParams['axes.unicode_minus'] = False
        return True
    except:
        return False

HAS_FONT = init_fonts()

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

# ==========================================
# 區塊 2: 網頁雷達圖產生器 (直接渲染在網頁)
# ==========================================
def draw_web_radar_chart(labels, pr_scores):
    num_vars = len(labels)
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
st.markdown("請選擇考試項目並輸入座號，即可查詢個人成績與分析圖表。")

# 🔗 【請在此替換成您剛剛建立的「總目錄主控表」網址】
MASTER_MENU_URL = "https://docs.google.com/spreadsheets/d/1lp0F45BnLO0Hn2l47vJ7orawr0KfIaT_/edit#gid=0"

csv_menu_url = get_google_sheet_csv_url(MASTER_MENU_URL)

if csv_menu_url:
    try:
        # 1. 讀取總目錄
        df_menu = pd.read_csv(csv_menu_url)
        df_menu.columns = df_menu.columns.str.strip()
        exam_options = df_menu['考試名稱'].dropna().tolist()
        
        # 2. 選擇考試下拉選單
        selected_exam = st.selectbox("📅 請選擇考試項目：", exam_options)
        
        # 抓取對應考試的獨立檔案網址
        target_sheet_url = df_menu[df_menu['考試名稱'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
        target_csv_url = get_google_sheet_csv_url(target_sheet_url)
        
        if target_csv_url:
            # 3. 讀取該次考試的獨立檔案
            raw_df = pd.read_csv(target_csv_url)
            raw_df.columns = raw_df.columns.str.strip()
            
            # 分離學生與統計列
            raw_df['_is_student'] = pd.to_numeric(raw_df['座號'], errors='coerce').notna()
            df_students = raw_df[raw_df['_is_student']].copy()
            df_students['座號'] = df_students['座號'].astype(int)
            df_students = df_students[df_students['座號'] > 0]
            
            df_bottom = raw_df[~raw_df['_is_student']].copy()
            
            # 提取高標與平均
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
            
            # 學生座號選單
            student_seats = sorted(df_students['座號'].tolist())
            selected_seat = st.selectbox("🔢 請選擇您的座號：", student_seats)
            
            if selected_seat:
                student_row = df_students[df_students['座號'] == selected_seat].iloc[0]
                
                # 全自動科目偵測 (檢查這次考試檔案有哪些科目)
                available_subjects = [c for c in cols_to_extract if c in raw_df.columns]
                has_7_subjects = all(c in available_subjects for c in ['歷史', '地理', '公民'])
                
                if has_7_subjects:
                    display_subs = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                    radar_subs = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
                else:
                    display_subs = ['國文', '英文', '數學', '自然', '社會']
                    radar_subs = ['國文', '英文', '數學', '自然', '社會']
                
                # 計算各科個人 PR 值
                for sub in radar_subs:
                    if sub in df_students.columns:
                        df_students[sub] = pd.to_numeric(df_students[sub], errors='coerce').fillna(0)
                        df_students[f'{sub}_PR'] = df_students[sub].rank(pct=True) * 100
                
                # 重新鎖定該生更新後的資料
                updated_student_df = df_students[df_students['座號'] == selected_seat].iloc[0]
                
                st.markdown(### f"👤 座號 {selected_seat} 號 【{updated_student_df['姓名']}】 的成績報告")
                
                # 建立網頁左右雙欄排版
                col1, col2 = st.columns([1.1, 1.0])
                
                with col1:
                    st.write("**📊 各科成績明細：**")
                    for s in display_subs:
                        if s in updated_student_df:
                            score = float(updated_student_df[s])
                            stat = stats_dict.get(s, {})
                            avg = stat.get('平均', 0)
                            high = stat.get('高標', 0)
                            
                            # 格式化輸出
                            if s in ['歷史', '地理', '公民']:
                                st.text(f"  └ {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                            else:
                                st.text(f"  {s}：{score:.2f}  (班均: {avg:.2f} | 高標: {high:.2f})")
                    
                    st.markdown("---")
                    # 總分與排名資訊
                    total_score = float(updated_student_df['總分']) if '總分' in updated_student_df else 0
                    rank = int(updated_student_df['名次']) if '名次' in updated_student_df else 0
                    total_pr = float(updated_student_df['總PR']) if '總PR' in updated_student_df else 0
                    
                    st.metric(label="五科總分", value=f"{total_score:.2f}")
                    st.write(f"🏆 **班級名次**：第 {rank} 名")
                    st.write(f"📈 **總分 PR 值**：{total_pr:.2f}")
                    
                with col2:
                    # 直接在網頁渲染雷達圖
                    pr_scores = [updated_student_df[f'{s}_PR'] for s in radar_subs]
                    fig = draw_web_radar_chart(radar_subs, pr_scores)
                    st.pyplot(fig)
                    
    except Exception as e:
        st.error(f"❌ 系統讀取錯誤，請確認各段考檔案欄位是否與公版一致。錯誤訊息: {e}")
else:
    st.error("❌ 無法載入總目錄，請確認網址權限設定。")
