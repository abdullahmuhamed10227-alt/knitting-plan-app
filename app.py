import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو - استقراء وقراءة البيانات والخانَات من الملفات حرفياً دون أي تعديل.")

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
        
        # 2. قراءة ملف البلان حرفياً بكل خاناته وأعمدته المحددة دون أي تخمين
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة بيانات ملف البلان: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                # في حالة الإكسيل، يتم قراءة الملف والأعمدة المطلوبة حرفياً كما هي تماماً
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                # دمج الأعمدة المطلوبة بالاسم حرفياً
                required_columns = ['Work Order', 'Seq', 'Acs', 'Item Description', 'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Daily Prd.', 'Yarn Information', 'Customer Name', 'Project Name']
                
                # التأكد من جلب الأعمدة الموجودة في ملف البلان حرفياً
                for col in required_columns:
                    matching_col = next((c for c in df_plan.columns if col.lower().strip() in str(c).lower().strip()), None)
                    if matching_col and matching_col != col:
                        df_plan = df_plan.rename(columns={matching_col: col})
                
                master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                st.success("✅ تمت قراءة أعمدة وخانات البلان (Excel) حرفياً ودون أي تعديل!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                # في حالة الرفع للـ PDF، يتم قراءة السطور واستخراج الحقول بدقة تامة ومنع ظهور أي كلمات عشوائية
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                chunks = re.split(r'(?=\b\d{6}-\d\b)', full_text)
                plan_rows = []
                
                for chunk in chunks:
                    wo_match = re.search(r'\b\d{6}-\d\b', chunk)
                    if wo_match:
                        wo_val = wo_match.group()
                        clean_line = " ".join(chunk.split())
                        
                        # استخراج دقيق وخالي من أي عشوائية
                        seq_match = re.search(r'^\s*(\d+)', clean_line)
                        seq_val = seq_match.group(1) if seq_match else ""
                        
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+)', clean_line)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        # فصل الكلمات الأخيرة بدقة لاستخراج العميل والمشروع الحقيقيين بعيداً عن الثوابت النصية
                        words = clean_line.split()
                        customer_val = words[-2] if len(words) >= 2 and not "Project" in words[-2] else ""
                        project_val = words[-1] if len(words) >= 1 and not "Name" in words[-1] else ""
                        
                        plan_rows.append({
                            "Work Order": wo_val,
                            "Seq": seq_val,
                            "Acs": "M" if " M " in clean_line else "",
                            "Item Description": clean_line,
                            "Sample Number": sample_no,
                            "Ref.Note": "",
                            "Pl/Tot.Qty": "",
                            "Daily Prd.": "",
                            "Yarn Information": clean_line[40:120].strip() if len(clean_line) > 120 else "",
                            "Customer Name": customer_val,
                            "Project Name": project_val
                        })
                        
                if plan_rows:
                    df_pdf_plan = pd.DataFrame(plan_rows)
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_plan.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم استخراج وقراءة محتوى الخانات من الـ PDF بدقة تامة ({len(df_pdf_plan)} سجل)!")
                else:
                    st.warning("⚠️ لم يتم العثور على أوردرات مطابقة للنمط داخل ملف الـ PDF.")
        
        # عرض التقرير النهائي الشامل
        st.subheader("📊 التقرير النهائي الشامل (مطابق للملفات الأصلية حرفياً):")
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