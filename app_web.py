import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import traceback

# ==========================================
# 1. 網頁頁面配置
# ==========================================
st.set_page_config(page_title="補考自動化神器-頂規網頁版", page_icon="🏫", layout="wide")

st.title("🏫 試務組-全校補考自動化神器 (Web 頂規大滿貫版)")
st.info("💡 網頁版優點：不用擔心檔案沒關閉導致當機，處理完直接點擊按鈕下載報表！")

# ==========================================
# 2. 輔助功能定義
# ==========================================
def get_str_col(df, keywords):
    if isinstance(keywords, str): keywords = [keywords]
    for kw in keywords:
        for i, col in enumerate(df.columns):
            if kw == str(col).strip():
                return df.iloc[:, i].fillna("").astype(str).str.strip()
    for kw in keywords:
        for i, col in enumerate(df.columns):
            if kw in str(col):
                return df.iloc[:, i].fillna("").astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index)

def grade_to_chinese(text):
    t = str(text)
    if '一' in t: return '一'
    if '二' in t: return '二'
    if '三' in t: return '三'
    if '1' in t or '１' in t: return '一'
    if '2' in t or '２' in t: return '二'
    if '3' in t or '３' in t: return '三'
    return "未知"

def natural_sort_key(s):
    return tuple(int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s)))

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==========================================
# 3. 介面佈局：功能選單
# ==========================================
st.divider()
col_files, col_opts = st.columns([1, 1], gap="large")

with col_files:
    st.subheader("📂 第一步：上傳原始資料")
    file_target = st.file_uploader("1️⃣ 補考名單.xlsx", type=['xlsx'])
    file_short = st.file_uploader("2️⃣ 科目簡稱.xlsx", type=['xlsx'])
    file_exam = st.file_uploader("3️⃣ 科目對照表.xlsx", type=['xlsx'])
    file_teacher = st.file_uploader("4️⃣ 監考教師及時間.xlsx", type=['xlsx'])

with col_opts:
    st.subheader("⚙️ 第二步：考場容量與分流設定")
    
    # 這裡就是你提到的關鍵選項！
    zhiyong_cap = st.radio(
        "📍 致用樓4樓會議室 人數上限：",
        [136, 148],
        index=0,
        horizontal=True,
        help="選擇 136 人可免去搬動桌椅程序。"
    )
    
    st.write("") # 間隔
    
    separate_mode = st.toggle(
        "🔥 開啟【多科與單科嚴格分流】功能",
        value=False,
        help="勾選後：致用樓僅排「多科」學生，剩下的空位不補滿。「單科」直接排圖書館。"
    )
    if separate_mode:
        st.warning("⚠️ 目前處於【分流模式】：多科排致用樓，單科排圖書館。")
    else:
        st.success("✅ 目前處於【滿額模式】：依考科數由多到少，自動填滿所有場地。")

# ==========================================
# 4. 執行與運算
# ==========================================
st.divider()

if st.button("🚀 開始智慧排考運算", type="primary", use_container_width=True):
    if not all([file_target, file_short, file_exam, file_teacher]):
        st.error("🚨 錯誤：請確認上方【4個檔案】皆已上傳完畢！")
    else:
        with st.spinner("系統正在執行：智慧分流、交叉打散、收卷動線最佳化..."):
            try:
                # 讀取 Excel
                df_short_map = pd.read_excel(file_short)
                df_exam_map = pd.read_excel(file_exam)
                df_target = pd.read_excel(file_target)
                df_teacher = pd.read_excel(file_teacher)

                LIBRARY_MAX = 98
                COMPUTER_MAX = 37

                # --- 階段一：基本資料處理 ---
                df_target['姓名'] = get_str_col(df_target, ['姓名', '學生姓名'])
                col_opencourse = get_str_col(df_target, ['開課班'])
                col_homeroom = get_str_col(df_target, ['班級', '原班級'])
                if col_opencourse.eq("").all(): col_opencourse = col_homeroom

                df_target['學號'] = get_str_col(df_target, ['學號'])
                df_target['座號'] = get_str_col(df_target, ['座號'])
                df_target['科目'] = get_str_col(df_target, ['科目', '考科'])
                df_target['班級'] = col_homeroom
                df_target['年級'] = df_target['班級'].apply(grade_to_chinese)

                # 科目簡稱
                col_a_s, col_b_s = df_short_map.columns[0], df_short_map.columns[1]
                short_dict = dict(zip(df_short_map[col_a_s].astype(str).str.strip(), df_short_map[col_b_s].astype(str).str.strip()))
                df_target['科目簡稱'] = df_target['科目'].map(short_dict).fillna("")

                # 試卷編號
                ex_cls = get_str_col(df_exam_map, ['班級', '開課班'])
                ex_sub = get_str_col(df_exam_map, ['科目', '考科'])
                ex_pap = get_str_col(df_exam_map, ['試卷編號', '代碼'])
                if ex_pap.eq("").all() and df_exam_map.shape[1] > 7: ex_pap = df_exam_map.iloc[:, 7].astype(str).str.strip()
                ex_dict = dict(zip(ex_cls + ex_sub, ex_pap))

                df_target['試卷編號'] = (col_opencourse + df_target['科目']).map(ex_dict).fillna((col_homeroom + df_target['科目']).map(ex_dict)).fillna("")
                df_target['試卷編號'] = df_target['試卷編號'].apply(lambda x: str(x).replace('.0','') if str(x).endswith('.0') else str(x))

                # 計算科目數
                df_temp = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first')
                v_counts = df_temp[df_temp['試卷編號'] != ""].groupby('學號').size()
                df_target['科目數目'] = df_target['學號'].map(v_counts).fillna(0).astype(int)

                # --- 智慧場地分配 ---
                df_students = df_target.drop_duplicates(subset=['學號']).copy()
                df_students = df_students[df_students['科目數目'] > 0]
                df_students = df_students.sort_values(by=['年級', '科目數目', '班級', '座號'], ascending=[True, False, True, True])

                venue_map = {}
                for gr, group in df_students.groupby('年級'):
                    if separate_mode:
                        multi = group[group['科目數目'] >= 2]
                        single = group[group['科目數目'] == 1]
                        m_v = (['致用樓四樓會議室'] * zhiyong_cap + ['圖書館三樓自修教室'] * LIBRARY_MAX + ['電腦教室401'] * COMPUTER_MAX)
                        m_ans = m_v[:len(multi)]
                        rem_l = max(0, LIBRARY_MAX - m_ans.count('圖書館三樓自修教室'))
                        rem_c = max(0, COMPUTER_MAX - m_ans.count('電腦教室401'))
                        s_v = (['圖書館三樓自修教室'] * rem_l + ['電腦教室401'] * rem_c)
                        s_ans = s_v[:len(single)]
                        if len(multi) > len(m_v): m_ans += ['超過容量未分配'] * (len(multi) - len(m_v))
                        if len(single) > len(s_v): s_ans += ['超過容量未分配'] * (len(single) - len(s_v))
                        for sid, v in zip(multi['學號'], m_ans): venue_map[sid] = v
                        for sid, v in zip(single['學號'], s_ans): venue_map[sid] = v
                    else:
                        vns = (['致用樓四樓會議室'] * zhiyong_cap + ['圖書館三樓自修教室'] * LIBRARY_MAX + ['電腦教室401'] * COMPUTER_MAX)
                        if len(group) > len(vns): vns += ['超過容量未分配'] * (len(group) - len(vns))
                        else: vns = vns[:len(group)]
                        for sid, v in zip(group['學號'], vns): venue_map[sid] = v

                df_target['場地'] = df_target['學號'].map(venue_map).fillna("")
                df_target.loc[df_target['試卷編號'] == "", '場地'] = "" 

                # 監考與時間
                df_teacher['比對年級'] = get_str_col(df_teacher, ['監考年級', '年級']).apply(grade_to_chinese)
                t_map = df_teacher.drop_duplicates(subset=['比對年級']).set_index('比對年級')[get_str_col(df_teacher, ['時間']).name].to_dict()
                df_target['時間2'] = df_target['年級'].map(t_map).fillna("")

                # --- 報表二處理 (標籤) ---
                df_label = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first').copy()
                df_label['單科標籤'] = df_label.apply(lambda r: f"{r['試卷編號']}{r['科目簡稱']}" if str(r['試卷編號']).strip() != "" else "", axis=1)
                
                df_grouped = df_label.groupby(['年級', '班級', '座號', '姓名', '科目數目', '場地'], dropna=False, as_index=False).agg({
                    '單科標籤': lambda x: '、'.join(sorted(dict.fromkeys([str(i) for i in x if str(i).strip() != ""]), key=natural_sort_key))
                })
                df_grouped = df_grouped.rename(columns={'科目數目': '個人考科', '場地': '地點', '單科標籤': '所有考科'})
                
                df_grouped['NumSeat'] = pd.to_numeric(df_grouped['座號'], errors='coerce').fillna(999)
                df_vld = df_grouped[df_grouped['地點'] != ""].copy()
                df_vld['G_W'] = df_vld['年級'].map(grade_weight).fillna(99)
                df_vld['L_W'] = df_vld['地點'].map(loc_weight).fillna(99)
                df_vld = df_vld.sort_values(by=['G_W', 'L_W', '個人考科', 'NumSeat', '班級'], ascending=[True, True, False, True, True])
                
                f_dfs = []
                for gr, loc_cap in [('一', current_targets), ('二', current_targets), ('三', current_targets)]:
                    for loc, cap in loc_cap.items():
                        sub = df_vld[(df_vld['地點'] == loc) & (df_vld['年級'] == gr)].copy()
                        if sub.empty: continue 
                        if len(sub) < cap: sub = pd.concat([sub, pd.DataFrame([{'地點': loc, '年級': gr}] * (cap - len(sub)))], ignore_index=True)
                        sub['序號'] = [f"{i+1:03d}" for i in range(len(sub))]
                        f_dfs.append(sub)
                
                df_rep2 = pd.concat(f_dfs, ignore_index=True) if f_dfs else pd.DataFrame(columns=['所有考科','班級','年級','座號','姓名','個人考科','地點','序號'])

                # --- 報表三處理 (考程) ---
                df_exam_sh = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first').copy()
                df_rep3_final = df_exam_sh.sort_values(by=['年級', '班級', '科目簡稱', '場地', '座號']).fillna("")

                # --- 報表四處理 (印卷) ---
                df_rep4 = df_target[df_target['試卷編號'] != ""].drop_duplicates(subset=['學號', '試卷編號']).groupby('試卷編號').size().reset_index(name='試卷數量')
                df_rep4['SortKey'] = df_rep4['試卷編號'].apply(natural_sort_key)
                df_rep4 = df_rep4.sort_values(by='SortKey').drop(columns=['SortKey'])

                # ==========================================
                # 5. 下載區
                # ==========================================
                st.balloons()
                st.success("🎊 處理成功！")
                
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    st.download_button("📄 下載：1.場地分配版", to_excel_bytes(df_target), "場地分配版.xlsx", "application/vnd.ms-excel", use_container_width=True)
                    st.download_button("🖨️ 下載：2.排座標籤", to_excel_bytes(df_rep2[['所有考科','班級','年級','座號','姓名','個人考科','地點','序號']]), "報表二_排座標籤.xlsx", "application/vnd.ms-excel", use_container_width=True)
                with d_col2:
                    st.download_button("📋 下載：3.考程匯整表", to_excel_bytes(df_rep3_final), "全校補考考程匯整表.xlsx", "application/vnd.ms-excel", use_container_width=True)
                    st.download_button("📝 下載：4.試卷印製表", to_excel_bytes(df_rep4), "報表四_試卷印製表.xlsx", "application/vnd.ms-excel", use_container_width=True)

            except Exception as e:
                st.error("🚨 發生錯誤！請檢查檔案格式是否正確。")
                st.expander("查看錯誤日誌").code(traceback.format_exc())

# 常數定義
grade_weight = {'一': 1, '二': 2, '三': 3}
loc_weight = {'致用樓四樓會議室': 1, '圖書館三樓自修教室': 2, '電腦教室401': 3}
current_targets = {"致用樓四樓會議室": zhiyong_cap, "圖書館三樓自修教室": 98, "電腦教室401": 37}
