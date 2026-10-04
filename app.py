import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو واستخراج وتحليل كافة تفاصيل ملفات البلان بدقة تامة.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك بالكامل
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # معالجة الخلايا المدمجة للأعمدة الإجمالية
        merge_columns_to_fill = ['ERP ROLLS', 'Roll / Mc  ERP', 'Unnamed: 21', 'Work ORDER', 'type Qualities', 'CUSTMER']
        for col in merge_columns_to_fill:
            if col in df_track.columns:
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة وتحليل ملف البلان واستخراج البيانات داخل الأعمدة بدقة حرفية
        if uploaded_plan is not None:
            st.info(f"📁 جاري تحليل واستخراج تفاصيل البلان من: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                st.success("✅ تمت إضافة أعمدة وبيانات البلان (Excel) بدقة تامة!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                # تقطيع النص بناءً على أوردرات التشغيل (مثال: 485419-1)
                chunks = re.split(r'(?=\b\d{6}-\d\b)', full_text)
                plan_rows = []
                
                for chunk in chunks:
                    wo_match = re.search(r'\b\d{6}-\d\b', chunk)
                    if wo_match:
                        wo_val = wo_match.group()
                        
                        # استخراج التسلسل (Seq) الذي يسبق رقم الأوردر غالباً
                        seq_match = re.search(r'(\d+)\s+' + wo_val, chunk)
                        seq_val = seq_match.group(1) if seq_match else ""
                        
                        # استخراج رقم العينة (Sample No) الذي يبدأ بـ I-N أو i-n
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+)', chunk)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        # استخراج معلومات الخيوط والغزل (Yarn Information)
                        yarn_match = re.search(r'(Ne\s+\d+/\d+[^|\n]+|OPENEND[^|\n]+|PENYE[^|\n]+)', chunk, re.IGNORECASE)
                        yarn_info = yarn_match.group(1).strip() if yarn_match else ""
                        
                        # استخراج اسم العميل والمشروع من نهاية النص الخاص بالأوردر
                        lines_in_chunk = [l.strip() for l in chunk.split('\n') if l.strip()]
                        customer_name = ""
                        project_name = ""
                        if lines_in_chunk:
                            last_line = lines_in_chunk[-1]
                            tokens = last_line.split()
                            if len(tokens) >= 2:
                                customer_name = tokens[-2]
                                project_name = tokens[-1]
                            elif len(tokens) == 1:
                                customer_name = tokens[0]

                        plan_rows.append({
                            "Work Order": wo_val,
                            "Seq": seq_val,
                            "Acs": "M" if " M " in chunk or chunk.endswith(" M") else "",
                            "Item Description": lines_in_chunk[0] if lines_in_chunk else "",
                            "Sample Number": sample_no,
                            "Ref.Note": "",
                            "Pl/Tot.Qty": "",
                            "Daily Prd.": "",
                            "Yarn Information": yarn_info,
                            "Customer Name": customer_name,
                            "Project Name": project_name
                        })
                        
                if plan_rows:
                    df_pdf_plan = pd.DataFrame(plan_rows)
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_plan.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم استخراج وتعبئة تفاصيل البلان بدقة تامة داخل الأعمدة ({len(df_pdf_plan)} أوردر مستخرج)!")
                else:
                    st.warning("⚠️ لم يتم العثور على أوردرات مطابقة للنمط داخل ملف الـ PDF.")
        
        # عرض التقرير النهائي الشامل
        st.subheader("📊 التقرير النهائي الشامل بعد تحليل وتوزيع بيانات البلان:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        # زر التحميل المباشر
        output_filename = "Master_Knitting_Report.xlsx"
        master_df.to_excel(output_filename, index=False)
        
        with open(output_filename, "rb") as file:
            st.download_button(
                label="📥 تحميل التقرير النهائي الشامل (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) وملف البلان للبدء.")