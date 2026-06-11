import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import io
import re
from matplotlib import font_manager
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="成績大數據主控台", layout="centered", page_icon="🎓")

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
    st.session_state.passwords = {}
if 'teacher_class_input' not in st.session_state:
    st.session_state.teacher_class_input = '08'  # 預設為 08 班，您可以隨時在右上角更改

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
        if "export?format=csv" in url:
            return url
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            gid_match = re.search(r'gid=([0-9]+)', url)
            gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
            return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv{gid_param}"
    except:
        pass
    return None

@st.cache_data(ttl=60)
def load_data(url):
    csv_url = get_google_sheet_csv_url(url)
    if not csv_url:
        return None
    try:
        df = pd.read_csv(csv_url)
        df.columns = df.columns.astype(str).str.strip()
        # 清除重複的欄位 (例如因為平均欄位導致多出好幾個'社會')
        df = df.loc[:, ~df.columns.duplicated(keep='first')]

        if '英語' in df.columns:
            df = df.rename(columns={'英語': '英文'})
        if '學號' in df.columns and '座號' not in df.columns:
            df = df.rename(columns={'學號': '座號'})
        if '班級座號' in df.columns and '座號' not in df.columns:
            df = df.rename(columns={'班級座號': '座號'})

        # 自動尋找學號欄位
        if '座號' not in df.columns:
            id_col = None
            for col in df.columns:
                vals = pd.to_numeric(df[col], errors='coerce')
                if vals.notna().any() and ((vals >= 70000) & (vals <= 99999)).any():
                    id_col = col
                    break
            if not id_col:
                for col in df.columns:
                    vals = pd.to_numeric(df[col], errors='coerce')
                    if vals.notna().sum() >= 3:
                        id_col = col
                        break
            if id_col:
                df = df.rename(columns={id_col: '座號'})

        return df
    except:
        return None

def get_student_mask(df):
    if '座號' not in df.columns:
        return pd.Series([False] * len(df))
    vals = pd.to_numeric(df['座號'], errors='coerce')
    mask_5digit = (vals >= 70000) & (vals <= 99999)
    mask_class = (vals >= 1) & (vals <= 60) & (vals == vals.round())
    stat_keywords = ['高標', '平均', '均標', '低標', '總計', '人數', '班平']
    if '姓名' in df.columns:
        is_stat = df['姓名'].astype(str).str.contains('|'.join(stat_keywords), na=False)
    else:
        is_stat = pd.Series([False] * len(df))
    return (mask_5digit | mask_class) & vals.notna() & ~is_stat

def check_is_grade_2_or_3(exam_name):
    return any(keyword in str(exam_name) for keyword in ["二上", "二下", "三上", "三下"])

# ==========================================
# 區塊 3: 導師功能 (全校總平均 & 會考落點)
# ==========================================
def get_menu_cols(df_menu):
    exam_col = '考試檔案名稱或別名' if '考試檔案名稱或別名' in df_menu.columns else '考試名稱'
    url_col = '該次成績單的 Google 試算表網址' if '該次成績單的 Google 試算表網址' in df_menu.columns else df_menu.columns[1]
    return exam_col, url_col

def generate_semester_average_excel(df_menu):
    all_data = []
    exam_col, url_col = get_menu_cols(df_menu)

    for _, row in df_menu.iterrows():
        exam_name = str(row.get(exam_col, '')).strip()
        file_url = str(row.get(url_col, '')).strip()
        if not file_url or file_url == 'nan':
            continue

        df_exam = load_data(file_url)
        if df_exam is None or '座號' not in df_exam.columns:
            continue

        df_stud = df_exam[get_student_mask(df_exam)].copy()
        is_grade_2_3 = check_is_grade_2_or_3(exam_name)

        for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
            if col not in df_stud.columns:
                df_stud[col] = np.nan
            else:
                df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce')

        if is_grade_2_3:
            df_stud['社會_融合'] = df_stud[['歷史', '地理', '公民']].mean(axis=1)
            df_stud['社會_融合'] = df_stud['社會_融合'].fillna(df_stud['社會'])
        else:
            df_stud['社會_融合'] = df_stud['社會']

        for _, s_row in df_stud.iterrows():
            seat_val = pd.to_numeric(s_row.get('座號'), errors='coerce')
            if pd.isna(seat_val): continue
            
            all_data.append({
                '座號': int(seat_val),
                '姓名': str(s_row.get('姓名', '-')).strip() if pd.notna(s_row.get('姓名')) else "-",
                '國文': s_row.get('國文', np.nan),
                '英文': s_row.get('英文', np.nan),
                '數學': s_row.get('數學', np.nan),
                '自然': s_row.get('自然', np.nan),
                '社會_融合': s_row.get('社會_融合', np.nan),
                '歷史': s_row.get('歷史', np.nan) if is_grade_2_3 else s_row.get('社會', np.nan),
                '地理': s_row.get('地理', np.nan) if is_grade_2_3 else s_row.get('社會', np.nan),
                '公民': s_row.get('公民', np.nan) if is_grade_2_3 else s_row.get('社會', np.nan),
            })

    if not all_data: return None

    df_all = pd.DataFrame(all_data)
    df_mean = df_all.groupby(['座號', '姓名']).mean(numeric_only=True).reset_index()

    sub5 = ['國文', '英文', '數學', '自然', '社會_融合']
    for col in sub5:
        if col not in df_mean.columns: df_mean[col] = np.nan

    valid_count = df_mean[sub5].notna().sum(axis=1)
    df_mean['五科總平均'] = df_mean[sub5].sum(axis=1, min_count=1) / valid_count.replace(0, np.nan)
    df_mean = df_mean.sort_values('座號')

    wb = Workbook()
    ws = wb.active
    ws.title = "全校學期總平均"
    headers = ["班級座號(學號)", "姓名", "國文平均", "英文平均", "數學平均", "自然平均",
               "社會(融合)平均", "歷史平均", "地理平均", "公民平均", "五科總和平均"]
    ws.append(headers)

    for _, r in df_mean.iterrows():
        def fmt(v):
            return round(v, 2) if pd.notna(v) else ""
        ws.append([
            r.get('座號', ''), r.get('姓名', ''),
            fmt(r.get('國文')), fmt(r.get('英文')), fmt(r.get('數學')),
            fmt(r.get('自然')), fmt(r.get('社會_融合')),
            fmt(r.get('歷史')), fmt(r.get('地理')), fmt(r.get('公民')),
            fmt(r.get('五科總平均')),
        ])

    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=11):
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1: cell.font = Font(bold=True)

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 14

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def pr_to_level(pr):
    if pd.isna(pr): return ""
    if pr >= 91: return "A++"
    elif pr >= 85: return "A+"
    elif pr >= 75: return "A"
    elif pr >= 60: return "B++"
    elif pr >= 45: return "B+"
    elif pr >= 15: return "B"
    else: return "C"

# 🌟 會考落點引擎：精準比對 2 碼或 3 碼，且姓名空白絕不閃退
def get_class_pr_data(df_menu, target_class):
    all_data = []
    exam_col, url_col = get_menu_cols(df_menu)
    target_class = str(target_class).strip() # 例如 '08' 或 '708'

    for _, row in df_menu.iterrows():
        exam_name = str(row.get(exam_col, '')).strip()
        file_url = str(row.get(url_col, '')).strip()
        if not file_url or file_url == 'nan': continue

        df_exam = load_data(file_url)
        if df_exam is None or '座號' not in df_exam.columns: continue

        df_stud = df_exam[get_student_mask(df_exam)].copy()
        if df_stud.empty: continue

        is_grade_2_3 = check_is_grade_2_or_3(exam_name)

        for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
            if col not in df_stud.columns: df_stud[col] = np.nan
            else: df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce')

        if is_grade_2_3:
            df_stud['社會_融合'] = df_stud[['歷史', '地理', '公民']].mean(axis=1)
            df_stud['社會_融合'] = df_stud['社會_融合'].fillna(df_stud['社會'])
        else:
            df_stud['社會_融合'] = df_stud['社會']

        # 計算全校 PR 值
        for sub in ['國文', '英文', '數學', '自然', '社會_融合']:
            col_pr = f'{sub}_PR'
            df_stud[col_pr] = np.nan
            valid_mask = df_stud[sub].notna()
            if valid_mask.sum() > 0:
                df_stud.loc[valid_mask, col_pr] = df_stud.loc[valid_mask, sub].rank(pct=True) * 100

        for _, s_row in df_stud.iterrows():
            seat_val = pd.to_numeric(s_row.get('座號'), errors='coerce')
            if pd.isna(seat_val): continue

            # 確保姓名為空時，塞入 "-"，避免被 groupby 放棄
            name_val = str(s_row.get('姓名', '-')).strip()
            if name_val == 'nan' or not name_val:
                name_val = "-"

            full_id_str = str(int(seat_val))
            if len(full_id_str) >= 4:
                class_part = full_id_str[:-2]  # 前面班號 (如 708)
                seat_id = int(full_id_str[-2:]) # 後面座號 (如 01)
                
                # 🌟 精準比對：輸入 08 可匹配 708、808；輸入 708 則僅匹配 708
                if not (class_part == target_class or class_part.endswith(target_class)):
                    continue
                student_key = full_id_str
            else:
                seat_id = int(seat_val)
                student_key = f"{target_class}_{seat_id:02d}"

            all_data.append({
                '學生識別碼': student_key,
                '座號': seat_id,
                '姓名': name_val,
                '國文PR': s_row.get('國文_PR', np.nan),
                '英文PR': s_row.get('英文_PR', np.nan),
                '數學PR': s_row.get('數學_PR', np.nan),
                '自然PR': s_row.get('自然_PR', np.nan),
                '社會PR': s_row.get('社會_融合_PR', np.nan),
                '考試名稱': exam_name,
            })

    if not all_data: return pd.DataFrame()

    df_all = pd.DataFrame(all_data)
    # dropna=False 保證空值不會害資料整列消失
    df_mean = df_all.groupby(['學生識別碼', '座號', '姓名'], dropna=False)[
        ['國文PR', '英文PR', '數學PR', '自然PR', '社會PR']
    ].mean().reset_index()

    df_mean['五科總PR平均'] = df_mean[['國文PR', '英文PR', '數學PR', '自然PR', '社會PR']].mean(axis=1)
    df_mean['國文預估'] = df_mean['國文PR'].apply(pr_to_level)
    df_mean['英文預估'] = df_mean['英文PR'].apply(pr_to_level)
    df_mean['數學預估'] = df_mean['數學PR'].apply(pr_to_level)
    df_mean['自然預估'] = df_mean['自然PR'].apply(pr_to_level)
    df_mean['社會預估'] = df_mean['社會PR'].apply(pr_to_level)
    df_mean['綜合預估'] = df_mean['五科總PR平均'].apply(pr_to_level)
    df_mean = df_mean.sort_values('學生識別碼').reset_index(drop=True)
    return df_mean

def generate_pr_report_excel(df_menu, target_class):
    df_mean = get_class_pr_data(df_menu, target_class)
    if df_mean.empty: return None, None

    wb = Workbook()
    ws = wb.active
    ws.title = "會考預估落點"
    headers = ["班級座號(學號)", "座號", "姓名",
               "國文PR", "國文預估", "英文PR", "英文預估",
               "數學PR", "數學預估", "自然PR", "自然預估",
               "社會PR", "社會預估", "五科平均PR", "綜合預估等級"]
    ws.append(headers)

    for _, r in df_mean.iterrows():
        def fmt(v):
            return round(float(v), 2) if pd.notna(v) else ""
        ws.append([
            r.get('學生識別碼', ''), r.get('座號', ''), r.get('姓名', ''),
            fmt(r.get('國文PR')), r.get('國文預估', ''),
            fmt(r.get('英文PR')), r.get('英文預估', ''),
            fmt(r.get('數學PR')), r.get('數學預估', ''),
            fmt(r.get('自然PR')), r.get('自然預估', ''),
            fmt(r.get('社會PR')), r.get('社會預估', ''),
            fmt(r.get('五科總PR平均')), r.get('綜合預估', ''),
        ])

    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    level_colors = {'A++': 'C6EFCE', 'A+': 'C6EFCE', 'A': 'C6EFCE',
                    'B++': 'FFEB9C', 'B+': 'FFEB9C', 'B': 'FFEB9C', 'C': 'FFC7CE'}
    from openpyxl.styles import PatternFill
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.value in level_colors:
                cell.fill = PatternFill("solid", fgColor=level_colors[cell.value])
    for cell in ws[1]:
        cell.border = thin
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(bold=True)
    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 13

    out = io.BytesIO()
    wb.save(out)
    return df_mean, out.getvalue()

# ==========================================
# PDF 輸出：全班會考預估落點報告
# ==========================================
GSAT_REFERENCE = [
    ["等級", "PR 範圍", "意義說明"],
    ["A++", "PR 91~99", "全國前 9%，高度精熟"],
    ["A+",  "PR 85~90", "全國前 15%，精熟"],
    ["A",   "PR 75~84", "全國前 25%，達精熟"],
    ["B++", "PR 60~74", "全國前 40%，基礎偏上"],
    ["B+",  "PR 45~59", "全國前 55%，基礎"],
    ["B",   "PR 15~44", "基礎程度"],
    ["C",   "PR 14 以下", "待加強"],
]

def generate_pr_pdf(df_mean, class_name=""):
    try:
        font_path = "NotoSansTC-Regular.ttf"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont("NotoSans", font_path))
            font_name = "NotoSans"
        else: font_name = "Helvetica"

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('title', fontName=font_name, fontSize=16, spaceAfter=6, alignment=1, textColor=colors.HexColor('#1a3a5c'))
        sub_style  = ParagraphStyle('sub', fontName=font_name, fontSize=10, spaceAfter=4, alignment=1, textColor=colors.grey)
        head_style = ParagraphStyle('head', fontName=font_name, fontSize=12, spaceBefore=12, spaceAfter=4, textColor=colors.HexColor('#1a3a5c'))
        note_style = ParagraphStyle('note', fontName=font_name, fontSize=8, textColor=colors.grey)

        story = []
        story.append(Paragraph(f"{'  ' + class_name + '  ' if class_name else ''}會考預估落點分析報告", title_style))
        story.append(Paragraph("根據歷次段考 PR 值加權平均推估・僅供參考", sub_style))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1a3a5c')))
        story.append(Spacer(1, 0.3*cm))

        story.append(Paragraph("▌ 各科預估落點總表", head_style))

        pr_cols = ['國文PR', '英文PR', '數學PR', '自然PR', '社會PR']
        lv_cols = ['國文預估', '英文預估', '數學預估', '自然預估', '社會預估']
        table_data = [["座號", "姓名", "國文", "英文", "數學", "自然", "社會", "綜合\n預估", "五科\n平均PR"]]
        
        for _, r in df_mean.iterrows():
            def fmt_pr(v, lv): return f"{v:.0f}\n({lv})" if pd.notna(v) else "-"
            row_data = [str(int(r['座號'])) if pd.notna(r['座號']) else '-', str(r.get('姓名', '-'))]
            for pr_c, lv_c in zip(pr_cols, lv_cols):
                row_data.append(fmt_pr(r.get(pr_c), r.get(lv_c, '')))
            row_data.append(str(r.get('綜合預估', '-')))
            avg_pr = r.get('五科總PR平均')
            row_data.append(f"{avg_pr:.1f}" if pd.notna(avg_pr) else '-')
            table_data.append(row_data)

        col_widths = [1.2*cm, 2.2*cm, 2.1*cm, 2.1*cm, 2.1*cm, 2.1*cm, 2.1*cm, 1.6*cm, 1.8*cm]
        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        ts = TableStyle([
            ('FONTNAME',    (0,0), (-1,-1), font_name),
            ('FONTSIZE',    (0,0), (-1,-1), 8),
            ('FONTSIZE',    (0,0), (-1, 0), 9),
            ('BACKGROUND',  (0,0), (-1, 0), colors.HexColor('#1a3a5c')),
            ('TEXTCOLOR',   (0,0), (-1, 0), colors.white),
            ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
            ('GRID',        (0,0), (-1,-1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0f4f8')]),
            ('TOPPADDING',  (0,0), (-1,-1), 4),
            ('BOTTOMPADDING',(0,0), (-1,-1), 4),
        ])
        
        level_bg = {'A++': colors.HexColor('#C6EFCE'), 'A+': colors.HexColor('#C6EFCE'), 'A': colors.HexColor('#C6EFCE'),
                    'B++': colors.HexColor('#FFEB9C'), 'B+': colors.HexColor('#FFEB9C'), 'B': colors.HexColor('#FFEB9C'),
                    'C':   colors.HexColor('#FFC7CE')}
        
        for i, r in enumerate(df_mean.itertuples(), start=1):
            for ci, lv_col in zip([2, 3, 4, 5, 6, 7], lv_cols + ['綜合預估']):
                lv = getattr(r, lv_col.replace(' ', '_'), '')
                if lv in level_bg:
                    ts.add('BACKGROUND', (ci, i), (ci, i), level_bg[lv])
                    ts.add('TEXTCOLOR', (ci, i), (ci, i), colors.HexColor('#1a3a5c'))

        t.setStyle(ts)
        story.append(t)
        story.append(Spacer(1, 0.5*cm))

        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
        story.append(Paragraph("▌ 全國教育會考等級對照說明", head_style))
        ref_table = Table(GSAT_REFERENCE, colWidths=[2*cm, 3*cm, 9*cm])
        ref_ts = TableStyle([
            ('FONTNAME',   (0,0), (-1,-1), font_name),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#1a3a5c')),
            ('TEXTCOLOR',  (0,0), (-1, 0), colors.white),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.grey),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ])
        for lv, ri in {'A++':1,'A+':2,'A':3,'B++':4,'B+':5,'B':6,'C':7}.items():
            if lv in level_bg: ref_ts.add('BACKGROUND', (0, ri), (-1, ri), level_bg[lv])
        ref_table.setStyle(ref_ts)
        story.append(ref_table)

        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph("※ 本報告以歷次段考在全校的 PR 值為基礎，不受各次考試難易度影響，是預測會考落點最穩定的指標。", note_style))
        doc.build(story)
        return buf.getvalue()
    except Exception as e: return None

# ==========================================
# 區塊 4: 雷達圖 & 歷史數據
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
    ax.plot(angles, [50] * len(angles), color='#C0504D', linewidth=1.5, linestyle='--', label='PR 50 (中位數)')
    ax.plot(angles, plot_scores, color='#4F81BD', linewidth=2, linestyle='solid', label='個人表現')
    ax.fill(angles, plot_scores, color='#4F81BD', alpha=0.35)
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fontsize=9, ncol=2)
    return fig

def fetch_student_history(df_menu, seat_num):
    history_records = []
    exam_col, url_col = get_menu_cols(df_menu)

    for _, row in df_menu.iterrows():
        exam_name = str(row.get(exam_col, '')).strip()
        file_url = str(row.get(url_col, '')).strip()
        if not file_url or file_url == 'nan': continue

        df_exam = load_data(file_url)
        if df_exam is None or '座號' not in df_exam.columns: continue

        df_stud = df_exam[get_student_mask(df_exam)].copy()
        df_stud['座號'] = pd.to_numeric(df_stud['座號'], errors='coerce').astype(int)
        is_grade_2_3 = check_is_grade_2_or_3(exam_name)

        for col in ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']:
            if col not in df_stud.columns: df_stud[col] = np.nan
            df_stud[col] = pd.to_numeric(df_stud[col], errors='coerce')

        if is_grade_2_3:
            df_stud['社會_融合'] = df_stud[['歷史', '地理', '公民']].mean(axis=1).fillna(df_stud['社會'])
        else:
            df_stud['社會_融合'] = df_stud['社會']

        for sub in ['國文', '英文', '數學', '自然', '社會_融合']:
            col_pr = f'{sub}_PR'
            df_stud[col_pr] = np.nan
            valid_mask = df_stud[sub].notna()
            if valid_mask.sum() > 0:
                df_stud.loc[valid_mask, col_pr] = df_stud.loc[valid_mask, sub].rank(pct=True) * 100

        score_cols = ['國文', '英文', '數學', '自然', '社會_融合']
        df_stud['總分_calc'] = df_stud[score_cols].sum(axis=1, min_count=1)
        df_stud['名次_calc'] = df_stud['總分_calc'].rank(ascending=False, method='min')

        match = df_stud[df_stud['座號'] == seat_num]
        if not match.empty:
            s_row = match.iloc[0]
            history_records.append({
                "考試名稱": exam_name,
                "國文": s_row.get('國文_PR', np.nan), "英文": s_row.get('英文_PR', np.nan),
                "數學": s_row.get('數學_PR', np.nan), "自然": s_row.get('自然_PR', np.nan),
                "社會(含領域融合)": s_row.get('社會_融合_PR', np.nan),
                "總分": pd.to_numeric(s_row.get('總分_calc'), errors='coerce'),
                "名次": pd.to_numeric(s_row.get('名次_calc'), errors='coerce'),
            })
    return pd.DataFrame(history_records)

# ==========================================
# 區塊 5: 登入介面
# ==========================================
if not st.session_state.logged_in:
    st.title("🔐 專屬成績系統 - 登入")
    st.markdown("---")

    login_role = st.radio("請選擇您的身分：", ["導師", "學生", "家長"])
    if login_role == "導師":
        username = st.text_input("導師帳號：")
    else:
        username = st.text_input("請輸入班級座號 (5位數，例如 70101)：")

    password = st.text_input("請輸入密碼 (預設密碼為座號)：", type="password")

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
                st.error("❌ 學生或家長帳號請輸入 5 位數純數字！")
    st.stop()

# ==========================================
# 區塊 6: 主介面
# ==========================================
col_title, col_logout = st.columns([0.65, 0.35])
with col_title:
    st.title("🎓 成績大數據主控台")
with col_logout:
    role_name = "導師" if st.session_state.role == "teacher" else f"{'學生' if st.session_state.role == 'student' else '家長'} ({st.session_state.user_id})"
    st.info(f"👤 {role_name}")
    
    # 🌟 導師專屬全域班級切換器：利用 Streamlit 原生 key 機制，絕不會失憶
    if st.session_state.role == "teacher":
        st.text_input("👨‍🏫 篩選顯示班級 (如 08 代表 8 班，708 代表 7年8班)：", key="teacher_class_input")
            
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

if df_menu is None:
    st.error("❌ 無法載入您的主控總表，請確認網址權限設定。")
    st.stop()

student_seats = []
df_students = pd.DataFrame()

if st.session_state.role == "teacher":
    with st.expander("📊 導師專屬：匯出全校學期各科總平均 Excel", expanded=False):
        st.write("系統將自動讀取主控台內的所有段考成績，計算全校每位學生的「各科跨學期平均」與「五科總平均」。")
        if st.button("🚀 結算並匯出全校總表", type="primary"):
            with st.spinner("正在結算各次段考數據中，請稍候..."):
                excel_data = generate_semester_average_excel(df_menu)
                if excel_data:
                    st.success("✅ 全校學期總平均計算完成！")
                    st.download_button(
                        label="📥 下載【全校學期總平均 Excel】",
                        data=excel_data,
                        file_name="全校學期總平均總表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 運算失敗，請確認主控台內是否有成績資料。")
    st.markdown("---")

exam_col, url_col = get_menu_cols(df_menu)
exam_options = df_menu[exam_col].dropna().tolist()

tab1, tab2, tab3 = st.tabs(["🔍 單次段考詳細成績", "📈 歷史軌跡動態分析", "🎯 會考落點預估 (PR分析)"])

with tab1:
    selected_exam = st.selectbox("📅 請選擇要查看的考試項目：", exam_options, key="student_exam")
    target_sheet_url = df_menu[df_menu[exam_col] == selected_exam][url_col].values[0]
    has_7_subjects = check_is_grade_2_or_3(selected_exam)
    raw_df = load_data(target_sheet_url)

    if raw_df is not None and '座號' in raw_df.columns:
        student_mask = get_student_mask(raw_df)
        df_students = raw_df[student_mask].copy()
        df_students['座號'] = pd.to_numeric(df_students['座號'], errors='coerce').astype(int)
        df_bottom = raw_df[~student_mask].copy()

        stats_dict = {}
        cols_to_extract = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']

        for col in cols_to_extract:
            if col not in df_students.columns: df_students[col] = np.nan
            df_students[col] = pd.to_numeric(df_students[col], errors='coerce')
            if col not in df_bottom.columns: df_bottom[col] = np.nan

        for _, brow in df_bottom.iterrows():
            seat_str = str(brow.get('座號', '')).strip()
            name_str = str(brow.get('姓名', '')).strip()
            label = seat_str + name_str
            for keyword, stat_key in [('高標', '高標'), ('平均', '平均'), ('均標', '平均')]:
                if keyword in label:
                    for c in cols_to_extract:
                        if c not in stats_dict: stats_dict[c] = {}
                        stats_dict[c][stat_key] = pd.to_numeric(brow.get(c), errors='coerce')

        if has_7_subjects:
            display_subs = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
            radar_subs = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
            df_students['社會_融合'] = df_students[['歷史', '地理', '公民']].mean(axis=1).fillna(df_students['社會'])
        else:
            display_subs = ['國文', '英文', '數學', '自然', '社會']
            radar_subs = ['國文', '英文', '數學', '自然', '社會']
            df_students['社會_融合'] = df_students['社會']

        score_cols = ['國文', '英文', '數學', '自然', '社會_融合']
        df_students['總分_calc'] = df_students[score_cols].sum(axis=1, min_count=1)
        df_students['總PR_calc'] = df_students['總分_calc'].rank(pct=True) * 100
        df_students['名次_calc'] = df_students['總分_calc'].rank(ascending=False, method='min')

        for sub in radar_subs:
            col_pr = f'{sub}_PR_current'
            df_students[col_pr] = np.nan
            valid_mask = df_students[sub].notna()
            if valid_mask.sum() > 0:
                df_students.loc[valid_mask, col_pr] = df_students.loc[valid_mask, sub].rank(pct=True) * 100

        df_students = df_students.sort_values(by='座號').reset_index(drop=True)
        all_student_seats = sorted(df_students[df_students['座號'] > 0]['座號'].tolist())

        if st.session_state.role == "teacher":
            target_class = st.session_state.teacher_class_input.strip()
            # 🌟 精準比對前三碼為該班級 (輸入 08 匹配 708/808，輸入 708 匹配 708)
            student_seats = [s for s in all_student_seats if str(s)[:-2].endswith(target_class) or str(s)[:-2] == target_class]
            if student_seats:
                selected_seat = st.selectbox(f"🔢 請選擇座號 (目前鎖定 {target_class} 班)：", student_seats, key="student_seat")
            else:
                st.warning(f"⚠️ 找不到 {target_class} 班的學生資料。")
                selected_seat = None
        else:
            student_seats = all_student_seats
            selected_seat = st.session_state.user_id

        if selected_seat and selected_seat in student_seats:
            stud_data = df_students[df_students['座號'] == selected_seat].iloc[0]
            st.markdown(f"### 👤 座號 {selected_seat} 號 【{stud_data.get('姓名', '-')}】 的成績報告")

            col1, col2 = st.columns([1.1, 1.0])
            with col1:
                st.write("**📊 各科成績明細：**")
                for s in display_subs:
                    if s in stud_data and pd.notna(stud_data[s]):
                        score = float(stud_data[s])
                        stat = stats_dict.get(s, {})
                        avg = float(stat.get('平均', 0)) if pd.notna(stat.get('平均')) else 0.0
                        high = float(stat.get('高標', 0)) if pd.notna(stat.get('高標')) else 0.0
                        indent = "  └ " if s in ['歷史', '地理', '公民'] else "  "
                        st.text(f"{indent}{s}：{score:.2f}  (全校平均: {avg:.2f} | 高標: {high:.2f})")

                st.markdown("---")
                total_score = float(stud_data.get('總分_calc', 0) or 0)
                rank = int(float(stud_data.get('名次_calc', 0) or 0))
                total_pr = float(stud_data.get('總PR_calc', 0) or 0)

                st.metric(label="五科總分", value=f"{total_score:.2f}")
                st.write(f"🏆 **全校排名**：第 {rank} 名")
                st.write(f"📈 **總分 PR 值**：{total_pr:.2f}")

            with col2:
                pr_scores = [stud_data.get(f'{s}_PR_current', 0) for s in radar_subs]
                fig = draw_web_radar_chart(radar_subs, pr_scores)
                st.pyplot(fig)
        else:
            st.warning("⚠️ 此班級或座號目前無成績資料。")
    else:
        st.warning("無法載入此考試資料。")

with tab2:
    st.markdown("### 📊 學生歷史跨學期趨勢追蹤")
    if not student_seats:
        st.info("請先確認您所在的班級有載入成績名單。")
    else:
        if st.session_state.role == "teacher":
            tc = st.session_state.teacher_class_input.strip()
            analysis_seat = st.selectbox(f"🔢 請選擇要分析的學生座號 (目前鎖定 {tc} 班)：", student_seats, key="analysis_seat")
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
                    if sub in df_history.columns:
                        ax_score.plot(df_history["考試名稱"], df_history[sub], marker='o', linewidth=2, label=f"{sub} PR")

                ax_score.set_title("各科 PR 值走勢圖 (跨學期全校能力追蹤)", fontsize=14, fontweight='bold')
                ax_score.set_ylabel("PR 值 (越大表現越優異)", fontsize=12)
                ax_score.set_ylim(0, 105)
                ax_score.grid(True, linestyle=':', alpha=0.6)
                ax_score.legend(loc='lower left', ncol=5, fontsize=9)
                plt.xticks(rotation=15)
                st.pyplot(fig_score)

                if "名次" in df_history.columns:
                    fig_rank, ax_rank = plt.subplots(figsize=(8, 3.5))
                    ax_rank.plot(df_history["考試名稱"], df_history["名次"], marker='s', color='#E67E22', linewidth=2.5, label="全校排名")
                    ax_rank.set_title("全校名次晉升軌跡 (曲線往上代表進步)", fontsize=14, fontweight='bold')
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

with tab3:
    st.markdown("### 🎯 國中教育會考 ─ PR 落點對照標準")
    st.info("平時段考分數會因為難易度而起伏，但 **PR 值 (百分等級)** 代表你在全校的相對排名位置，是預測會考落點最精準的數字！")

    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.success("**🏆 A 級 (精熟)**\n\n**A++**：PR 91 ～ 99\n**A+**：PR 85 ～ 90\n**A**：PR 75 ～ 84\n\n*(目標拿 A 者，各科請維持 PR 75 以上)*")
    with col_t2:
        st.warning("**💪 B 級 (基礎)**\n\n**B++**：PR 60 ～ 74\n**B+**：PR 45 ～ 59\n**B**：PR 15 ～ 44\n\n*(多數同學落點，拚進步的關鍵)*")
    with col_t3:
        st.error("**⚠️ C 級 (待加強)**\n\n**C**：PR 14 以下\n\n*(基礎觀念需要重新補強)*")

    st.markdown("---")

    if st.session_state.role == "teacher":
        tc = st.session_state.teacher_class_input.strip()
        st.markdown(f"#### 👨‍🏫 導師專屬：【{tc}】班 會考預估落點總表結算")
        
        if st.button(f"🚀 結算 {tc} 班會考預估落點", type="primary"):
            with st.spinner(f"正在全校資料庫中撈取符合 {tc} 班的學生，計算會考等級..."):
                df_pr, excel_pr = generate_pr_report_excel(df_menu, tc)
                if df_pr is not None and not df_pr.empty:
                    st.success(f"✅ 成功結算符合 {tc} 班共 {len(df_pr)} 位學生的落點！")
                    
                    st.download_button(
                        label=f"📥 下載【{tc}班會考預估落點 Excel】",
                        data=excel_pr,
                        file_name=f"班級{tc}_會考預估落點總表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    pdf_data = generate_pr_pdf(df_pr, class_name=tc)
                    if pdf_data:
                        st.download_button(
                            label=f"📄 下載【{tc}班會考預估落點 PDF】",
                            data=pdf_data,
                            file_name=f"班級{tc}_會考預估落點報告.pdf",
                            mime="application/pdf"
                        )
                    
                    st.dataframe(df_pr)
                else:
                    st.error(f"❌ 找不到包含代碼 '{tc}' 的班級資料。")
    else:
        st.markdown("#### 🎯 你的個人專屬預估落點")
        if st.button("🚀 分析我的會考落點", type="primary"):
            if st.session_state.user_id in student_seats:
                df_history = fetch_student_history(df_menu, st.session_state.user_id)
                if not df_history.empty:
                    my_pr = {
                        "國文": df_history["國文"].mean(),
                        "英文": df_history["英文"].mean(),
                        "數學": df_history["數學"].mean(),
                        "自然": df_history["自然"].mean(),
                        "社會": df_history["社會(含領域融合)"].mean() if "社會(含領域融合)" in df_history.columns else np.nan,
                    }
                    st.write("根據你歷次段考的 PR 值平均，為你預估的會考落點如下：")
                    cols = st.columns(5)
                    for i, (sub, pr) in enumerate(my_pr.items()):
                        if pd.notna(pr):
                            cols[i].metric(label=f"{sub} (平均 PR: {pr:.1f})", value=pr_to_level(pr))

                    total_pr = pd.Series(list(my_pr.values())).mean()
                    if pd.notna(total_pr):
                        st.info(f"✨ **你的五科綜合預估落點：{pr_to_level(total_pr)}** (總平均 PR: {total_pr:.1f})")
            else:
                st.warning("⚠️ 查無歷史資料，無法分析。")
