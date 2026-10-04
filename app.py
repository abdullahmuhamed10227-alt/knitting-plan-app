import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو واستخراج بيانات البلان وتوزيعها داخل أعمدتها المخصصة بدقة حرفية.")

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
        
        # 2. قراءة وتحليل ملف البلان وتوزيع كل قيمة في عمودها الصحيح
        if uploaded_plan is not None:
            st.info(f"📁 جاري تحليل بيانات البلان من: {uploaded_plan.name}...")
            
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
                
                # تقطيع النص بناءً على أوردرات التشغيل
                chunks = re.split(r'(?=\b\d{6}-\d\b)', full_text)
                plan_rows = []
                
                for chunk in chunks:
                    wo_match = re.search(r'\b\d{6}-\d\b', chunk)
                    if wo_match:
                        wo_val = wo_match.group()
                        cleaned_chunk = " ".join(chunk.split())
                        parts = cleaned_chunk.split()
                        
                        # 1. استخراج التسلسل (Seq)
                        seq_val = parts[0] if parts and parts[0].isdigit() else ""
                        
                        # 2. استخراج Acs
                        acs_val = parts[2] if len(parts) > 2 and len(parts[2]) <= 3 else "M"
                        
                        # 3. استخراج رقم العينة (Sample Number)
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+)', chunk)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        # 4. استخراج معلومات الغزل (Yarn Information)
                        yarn_match = re.search(r'(Ne\s*\d+/\d+[^|]+|OPENEND[^|]+|PENYE[^|]+)', chunk, re.IGNORECASE)
                        yarn_info = yarn_match.group(1).strip() if yarn_match else ""
                        
                        # 5. استخراج اسم العميل والمشروع من نهاية السطر
                        tokens = [t for t in parts if not re.match(r'^\d+([./]\d+)?$', t) and t != wo_val]
                        customer_name = tokens[-2] if len(tokens) >= 2 else (parts[-2] if len(parts) >= 2 else "")
                        project_name = tokens[-1] if len(tokens) >= 1 else (parts[-1] if len(parts) >= 1 else "")
                        
                        plan_rows.append({
                            "Work Order": wo_val,
                            "Seq": seq_val,
                            "Acs": acs_val,
                            "Item Description": cleaned_chunk[:100],
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
                    st.success(f"✅ تم استخراج وتعبئة بيانات البلان بدقة تامة داخل الأعمدة ({len(df_pdf_plan)} سجل)!")
                else:
                    st.warning("⚠️ لم يتم العثور على أوردرات مطابقة داخل ملف الـ PDF.")
        
        # عرض التقرير النهائي الشامل
        st.subheader("📊 التقرير النهائي الشامل بعد توزيع بيانات البلان بدقة:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        # زر التحميل
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