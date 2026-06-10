import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import io  # 🌟 新增：用於支援 Excel 記憶體內緩衝下載
from matplotlib import font_manager
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side

st.set_page_config(page_title="801 專屬成績大數據主控台", layout="centered", page_icon="🎓")

# ==========================================
# 區塊 1: 登入系統狀態初始化
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'role' not in st.session_state:
    st.session_state.role = None
if 'user_id' not in st.session_state:
    st.session_state.user_id = None
if 'passwords' not in st.session_state:
    st.session_state.passwords = {} # 暫存學生與家長更改的密碼

# ==========================================
# 區塊 2: 本地字體初始化
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

def check_is_grade_2_3(exam_name):
    if any(keyword in str(exam_name) for keyword in ["二上", "二下", "三上", "三下"]):
        return True
    return False

# ==========================================
# 🌟 新增：導師專屬 ─ 全班各科總平均 Excel 產生器
# ==========================================
def generate_class_summary_excel(df_menu):
    all_exam_data = []
    
    for _, row in df_menu.iterrows():
        exam_name = str(row['考試檔案名稱或別名'] if '考試檔案名稱或別名' in df_menu.columns else row['考試名稱']).strip()
        file_url = str(row['該次成績單的 Google 試算表網址']).strip()
        
        df_exam = load_data(file_url)
        if df_exam is not None:
            df_exam['_is_student'] = pd.to_numeric(df_exam['座號'], errors='coerce').notna()
            df_stud = df_exam[df_exam['_is_student']].copy()
            df_stud['座號'] = df_stud['座號'].astype(int)
            
            is_grade_2_3 = check_is_grade_2_3(exam_name)
            
            # 確保標準八科欄位皆存在並轉換為數值
            for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
                if col in df_stud.columns:
                    df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce').fillna(0)
                else:
                    df_stud[col] = 0.0
            
            # 若為一年級（無歷史地理公民獨立欄位），將社會分數同步至三科以利後續平均計算
            if not is_grade_2_3:
                df_stud['歷史'] = df_stud['社會']
                df_stud['地理'] = df_stud['社會']
                df_stud['公民'] = df_stud['社會']
                
            for _, s_row in df_stud.iterrows():
                all_exam_data.append({
                    "座號": int(s_row['座號']),
                    "姓名": str(s_row['姓名']).strip(),
                    "國文": s_row['國文'],
                    "英文": s_row['英文'],
                    "數學": s_row['數學'],
                    "自然": s_row['自然'],
                    "歷史": s_row['歷史'],
                    "地理': s_row['地理'],
                    "公民": s_row['公民']
                })
                
    if not all_exam_data:
        return None
        
    # 整合所有考試並計算每位學生的各科平均值
    df_all = pd.DataFrame(all_exam_data)
    df_mean = df_all.groupby(['座號', '姓名'])[['國文', '英文', '數學', '自然', '歷史', '地理', '公民']].mean().reset_index()
    
    # 重新融合社會科平均並計算五科總平均
    df_mean['社會(三科平均)'] = (df_mean['歷史'] + df_mean['地理'] + df_mean['公民']) / 3
    df_mean['五科總平均'] = df_mean['國文'] + df_mean['英文'] + df_mean['數學'] + df_mean['自然'] + df_mean['社會(三科平均)']
    df_mean = df_mean.sort_values(by='座號').reset_index(drop=True)
    
    # 建立 openpyxl 工作簿進行結構化排版
    wb = Workbook()
    ws = wb.active
    ws.title = "學期全班各科總平均表"
    
    # 表頭欄位
    headers = ["座號", "姓名", "國文平均", "英文平均", "數學平均", "自然平均", "歷史平均", "地理平均", "公民平均", "五科總平均"]
    ws.append(headers)
    
    # 寫入學生平均數據（四捨五入至小數點後兩位）
    for _, r in df_mean.iterrows():
        ws.append([
            r['座號'], r['姓名'],
            round(r['國文'], 2), round(r['英文'], 2), round(r['數學'], 2), round(r['自然'], 2),
            round(r['歷史'], 2), round(r['地理'], 2), round(r['公民'], 2), round(r['五科總平均'], 2)
        ])
        
    # 報表樣式美化
    f_header = Font(name="微軟正黑體", size=11, bold=True)
    f_data = Font(name="微軟正黑體", size=11)
    al_center = Alignment(horizontal="center", vertical="center")
    al_right = Alignment(horizontal="right", vertical="center")
    thin_side = Side(style='thin', color='BFBFBF')
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # 設定表頭樣式
    for cell in ws[1]:
        cell.font = f_header
        cell.alignment = al_center
        cell.border = cell_border
        
    # 設定資料列樣式
    for row in ws.iter_rows(min_row=2, max_row=len(df_mean) + 1, min_col=1, max_col=10):
        for cell in row:
            cell.font = f_data
            cell.border = cell_border
            if cell.column > 2:
                cell.alignment = al_right
            else:
                cell.alignment = al_center
                
    # 自動調整欄寬
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 區塊 3: 雷達圖產生器 & 歷史數據撈取
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
            
            is_grade_2_3 = check_is_grade_2_3(exam_name)
            
            for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
                if col not in df_stud.columns: df_stud[col] = 0
                df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce').fillna(0)
            
            if is_grade_2_3:
                df_stud['社會_融合'] = (df_stud['歷史'] + df_stud['地理'] + df_stud['公民']) / 3
            else:
                df_stud['社會_融合'] = df_stud['社會']
            
            for sub in ['國文', '英文', '數學', '自然', '社會_融合']:
                df_stud[f'{sub}_PR'] = df_stud[sub].rank(pct=True) * 100
            
            if '總分' not in df_stud.columns:
                df_stud['總分'] = df_stud['國文'] + df_stud['英文'] + df_stud['數學'] + df_stud['自然'] + df_stud['社會_融合']
            df_stud['總分'] = pd.to_numeric(df_stud['總分'], errors='coerce').fillna(0)
            
            if '名次' not in df_stud.columns:
                df_stud['名次'] = df_stud['總分'].rank(ascending=False, method='min')
            
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
# 區塊 4: 登入介面邏輯
# ==========================================
if not st.session_state.logged_in:
    st.title("🔐 801 專屬成績系統 - 登入")
    st.markdown("---")
    
    login_role = st.radio("請選擇您的身分：", ["導師", "學生", "家長"])
    
    if login_role == "導師":
        username = st.text_input("導師帳號：")
    else:
        username = st.text_input("請輸入座號：")
        
    password = st.text_input("請輸入密碼 (學生/家長預設密碼為座號)：", type="password")
    
    if st.button("🚀 登入系統", type="primary"):
        if login_role == "導師":
            if username == "yen" and password == "Axlc1982":
                st.session_state.logged_in = True
                st.session_state.role = "teacher"
                st.session_state.user_id = "yen"
                st.rerun()
            else:
                st.error("❌ 導師帳號或密碼錯誤！")
        else:
            try:
                seat_num = int(username)
                user_key = f"{login_role}_{seat_num}"
                expected_password = st.session_state.passwords.get(user_key, str(seat_num))
                
                if password == expected_password:
                    st.session_state.logged_in = True
                    st.session_state.role = "student" if login_role == "學生" else "parent"
                    st.session_state.user_id = seat_num
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤！")
            except ValueError:
                st.error("❌ 學生或家長帳號請輸入純數字座號！")
    st.stop()

# ==========================================
# 區塊 5: 系統主介面 (登入後)
# ==========================================
col_title, col_logout = st.columns([0.8, 0.2])
with col_title:
    st.title("🎓 801 專屬成績大數據主控台")
with col_logout:
    role_name = "導師" if st.session_state.role == "teacher" else f"{'學生' if st.session_state.role == 'student' else '家長'} ({st.session_state.user_id}號)"
    st.info(f"👤 {role_name}")
    if st.button("登出"):
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.user_id = None
        st.rerun()

if st.session_state.role != "teacher":
    with st.expander("🔑 修改個人密碼"):
        new_pass = st.text_input("請輸入新密碼：", type="password")
        confirm_pass = st.text_input("再次確認新密碼：", type="password")
        if st.button("確認修改"):
            if new_pass == confirm_pass and new_pass != "":
                user_key = f"{'學生' if st.session_state.role == 'student' else '家長'}_{st.session_state.user_id}"
                st.session_state.passwords[user_key] = new_pass
                st.success("✅ 密碼修改成功！下次請使用新密碼登入。")
            else:
                st.error("❌ 兩次輸入的密碼不一致，或密碼不能為空。")

st.markdown("---")

MASTER_MENU_URL = "https://docs.google.com/spreadsheets/d/1FL-orK8H_oDrLuDg1pAijJICjHzPxJLcPIEngF0wko8/edit?gid=0#gid=0"
df_menu = load_data(MASTER_MENU_URL)

if df_menu is not None:
    # 🌟 新增：導師專屬 ─ 全班跨歷史成績總平均下載功能區
    if st.session_state.role == "teacher":
        with st.expander("📊 導師專屬功能：全班跨學期各科總平均統計"):
            st.markdown("點擊下方按鈕後，系統將自動提取雲端主控總表登錄之所有段考成績，動態結算全班學生的跨學期個人總平均明細。")
            if st.button("🚀 開始計算全班總平均"):
                with st.spinner("正在讀取雲端各段考數據並進行交叉運算..."):
                    excel_summary = generate_class_summary_excel(df_menu)
                    if excel_summary:
                        st.success("✅ 全班各科總平均計算完成！")
                        st.download_button(
                            label="📥 下載全班各科學期總平均 Excel 表",
                            data=excel_summary,
                            file_name="801_全班學期各科總平均匯總表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("❌ 提取失敗，請檢查主控總表內的試算表連結與權限。")
        st.markdown("---")

    if '考試檔案名稱或別名' in df_menu.columns:
        exam_options = df_menu['考試檔案名稱或別名'].dropna().tolist()
    else:
        exam_options = df_menu['考試名稱'].dropna().tolist()
    
    tab1, tab2 = st.tabs(["🔍 單次段考詳細成績", "📈 歷史軌跡動態 analysis"])
    
    with tab1:
        selected_exam = st.selectbox("📅 請選擇要查看的考試項目：", exam_options, key="student_exam")
        if '考試檔案名稱或別名' in df_menu.columns:
            target_sheet_url = df_menu[df_menu['考試檔案名稱或別名'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
        else:
            target_sheet_url = df_menu[df_menu['考試名稱'] == selected_exam]['該次成績單的 Google 試算表網址'].values[0]
            
        has_7_subjects = check_is_grade_2_3(selected_exam)
        
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
                            stats_dict[c]['high'] = pd.to_numeric(row[c], errors='coerce')
                if '平均' in name or '均標' in name:
                    for c in cols_to_extract:
                        if c in row:
                            if c not in stats_dict: stats_dict[c] = {}
                            stats_dict[c]['average'] = pd.to_numeric(row[c], errors='coerce')
            
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
            
            if st.session_state.role == "teacher":
                selected_seat = st.selectbox("🔢 請選擇要查詢的座號：", student_seats, key="student_seat")
            else:
                selected_seat = st.session_state.user_id
            
            if selected_seat in student_seats:
                stud_data = df_students[df_students['座號'] == selected_seat].iloc[0]
                st.markdown(f"### 👤 座號 {selected_seat} 號 【{stud_data['姓名']}】 的成績報告")
                
                col1, col2 = st.columns([1.1, 1.0])
                
                with col1:
                    st.write("**📊 各科成績明細：**")
                    for s in display_subs:
                        if s in stud_data:
                            score = float(stud_data[s]) if pd.notna(stud_data[s]) else 0.0
                            stat = stats_dict.get(s, {})
                            avg = float(stat.get('average', 0)) if pd.notna(stat.get('average')) else 0.0
                            high = float(stat.get('high', 0)) if pd.notna(stat.get('high')) else 0.0
                            
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
            else:
                st.warning("⚠️ 此次考試中查無此座號成績資料。")

    with tab2:
        st.markdown("### 📊 學生歷史跨學期趨勢追蹤")
        if 'student_seats' in locals() and student_seats:
            if st.session_state.role == "teacher":
                analysis_seat = st.selectbox("🔢 請選擇要分析的學生座號：", student_seats, key="analysis_seat")
            else:
                analysis_seat = st.session_state.user_id
            
            if analysis_seat in student_seats:
                with st.spinner("正在為您統整歷史數據..."):
                    df_history = fetch_student_history(df_menu, analysis_seat)
                    
                if not df_history.empty:
                    stud_name = df_students[df_students['座號'] == analysis_seat]['姓名'].values[0]
                    st.markdown(f"#### 📈 {analysis_seat} 號 【{stud_name}】 學習軌跡綜合分析")
                    
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
                st.warning("⚠️ 查無此座號成績資料。")
else:
    st.error("❌ 無法載入您的主控總表，請確認網址權限設定。")
