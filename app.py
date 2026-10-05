import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الذكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الذكي والديناميكي")
st.markdown("نظام موحد ومستقل لقراءة ومعالجة ملفات التراك وخطط التشغيل مهما كانت الاختلافات في الأعمدة أو الماكينات.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة وتحليل ملف التراك الأساسي...")
        
        # قراءة أولية لاكتشاف صف الهيدر الصحيح في ملف التراك
        df_track_raw = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=None)
        header_track_idx = 1
        for idx, row in df_track_raw.iterrows():
            row_str = str(row.values).lower()
            if 'work order' in row_str or 'machine' in row_str or 'type qualities' in row_str:
                header_track_idx = idx
                break
                
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=header_track_idx)
        df_track = df_track.dropna(how='all')
        
        # تعبئة الخلايا المدمجة للأعمدة الرئيسية لتجنب أي قيم فارغة في التراك
        for col in df_track.columns:
            col_str = str(col).strip().upper()
            if any(k in col_str for k in ['ROLLS', 'ORDER', 'CUSTMER', 'QUALITIES']):
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة ومعالجة ملف البلان بشكل ديناميكي كامل
        df_plan = pd.DataFrame()
        if uploaded_plan is not None:
            st.info(f"📁 جاري تحليل ومعالجة ملف البلان: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan_raw = pd.read_excel(uploaded_plan, header=None)
                header_plan_idx = 2
                for idx, row in df_plan_raw.iterrows():
                    row_str = str(row.values).lower()
                    if 'work order' in row_str or 'seq' in row_str:
                        header_plan_idx = idx
                        break
                
                df_plan = pd.read_excel(uploaded_plan, header=header_plan_idx)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                # اكتشاف وعلاج عمود الماكينة المدمج (المليء بالفراغات تحت اسم الماكينة الأولى)
                mc_col = None
                for c in df_plan.columns:
                    c_low = str(c).strip().lower()
                    if 'unnamed: 1' in c_low or 'machine' in c_low or 'mc' in c_low or c_low == '1':
                        mc_col = c
                        break
                if not mc_col and len(df_plan.columns) > 1:
                    mc_col = df_plan.columns[1]  # الافتراضي الشائع في ملفات البلان
                    
                if mc_col:
                    df_plan = df_plan.rename(columns={mc_col: 'Detected_Machine'})
                    # أهم خطوة لملء ماكينات الأوردرات المتعددة تحت بعضها تلقائياً
                    df_plan['Detected_Machine'] = df_plan['Detected_Machine'].replace('', pd.NA).ffill()
                
                st.success(f"✅ تم تحليل ملف البلان بنجاح وعدد السجلات: {len(df_plan)}")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                lines = full_text.split('\n')
                plan_rows = []
                last_mc = ""
                
                for line in lines:
                    clean_line = " ".join(line.split())
                    if not clean_line:
                        continue
                        
                    wo_match = re.search(r'\b\d{6}-\d\b', clean_line)
                    machine_match = re.search(r'\b(M\d+|\d{3,4})\b', clean_line)
                    if machine_match:
                        last_mc = machine_match.group()
                        
                    if wo_match:
                        tokens = clean_line.split()
                        seq_val = ""
                        for token in tokens[:3]:
                            if token.isdigit() and int(token) < 1000:
                                seq_val = token
                                break
                                
                        plan_rows.append({
                            "Work_Order_Plan": wo_match.group(),
                            "Detected_Machine": last_mc,
                            "Seq_Plan": seq_val,
                            "Item_Description_Plan": clean_line
                        })
                if plan_rows:
                    df_plan = pd.DataFrame(plan_rows)

            # 3. توحيد مفاتيح الربط والدمج الذكي غير المدمر للبيانات
            if not df_plan.empty:
                # توحيد أسماء أعمدة البلان ديناميكياً
                for col in df_plan.columns:
                    col_s = str(col).strip().lower()
                    if 'work order' in col_s or col_s == 'work order':
                        df_plan = df_plan.rename(columns={col: 'Work_Order_Plan'})
                    elif 'seq' in col_s:
                        df_plan = df_plan.rename(columns={col: 'Seq_Plan'})

                # البحث عن أعمدة المطابقة في التراك
                track_wo = next((c for c in master_df.columns if 'work order' in str(c).lower() or 'order' in str(c).lower()), None)
                track_mc = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower()), None)
                track_seq = next((c for c in master_df.columns if 'seq' in str(c).lower()), None)
                
                if track_wo and 'Work_Order_Plan' in df_plan.columns:
                    master_df['Key_WO'] = master_df[track_wo].astype(str).str.strip().str.upper()
                    df_plan['Key_WO'] = df_plan['Work_Order_Plan'].astype(str).str.strip().str.upper()
                    
                    merge_keys = ['Key_WO']
                    
                    if track_mc and 'Detected_Machine' in df_plan.columns:
                        master_df['Key_MC'] = master_df[track_mc].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                        df_plan['Key_MC'] = df_plan['Detected_Machine'].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                        merge_keys.append('Key_MC')
                        
                    if track_seq and 'Seq_Plan' in df_plan.columns:
                        master_df['Key_Seq'] = master_df[track_seq].astype(str).str.strip().str.replace('.0', '', regex=False)
                        df_plan['Key_Seq'] = df_plan['Seq_Plan'].astype(str).str.strip().str.replace('.0', '', regex=False)
                        merge_keys.append('Key_Seq')

                    # إزالة التكرارات من البلان لضمان عدم مضاعفة الصفوف في التراك
                    df_plan = df_plan.drop_duplicates(subset=merge_keys, keep='first')

                    # إزالة أعمدة البلان القديمة المتداخلة من الماستر لو وجدت
                    for c in df_plan.columns:
                        if c in master_df.columns and c not in merge_keys:
                            master_df = master_df.drop(columns=[c])

                    # الدمج الذكي النهائي
                    master_df = pd.merge(master_df, df_plan, on=merge_keys, how='left')
                    
                    # تنظيف مفاتيح الربط المؤقتة
                    for k in ['Key_WO', 'Key_MC', 'Key_Seq']:
                        if k in master_df.columns:
                            master_df = master_df.drop(columns=[k])
                            
                    st.success("🔗 تم ربط البيانات واستخراجها بنجاح تام وفقاً لهيكل وتوزيع الملفات!")

        master_df = master_df.fillna("")

        st.subheader("📊 معاينة التقرير النهائي الموحد:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        output_filename = "Master_Knitting_System_Report.xlsx"
        master_df.to_excel(output_filename, index=False)
        
        with open(output_filename, "rb") as file:
            st.download_button(
                label="📥 تحميل التقرير النهائي (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء معالجة الملفات: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك الأساسي وملف البلان للبدء.")