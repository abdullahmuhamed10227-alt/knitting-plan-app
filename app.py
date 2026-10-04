import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام إدارة ومتابعة التريكو ودمج تفاصيل البلان داخل أعمدتها المخصصة بدقة حرفية.")

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
        
        # 2. قراءة وتحليل ملف البلان واستخراج البيانات داخل الأعمدة بدقة
        if uploaded_plan is not None:
            st.info(f"📁 جاري استخراج وتوزيع بيانات البلان من: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                st.success("✅ تمت اضافة أعمدة وبيانات البلان (Excel) بدقة تامة!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                plan_rows = []
                
                for page_idx, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    lines = text.split('\n')
                    for line in lines:
                        wo_match = re.search(r'\b\d{6}-\d\b', line)
                        if wo_match:
                            parts = line.split()
                            seq_val = parts[0] if parts and parts[0].isdigit() else ""
                            wo_val = wo_match.group()
                            
                            # استخراج المعلومات المتاحة وتوزيعها داخل الأعمدة المطلوبة بدقة
                            item_desc = " ".join(parts[3:12]) if len(parts) > 12 else line
                            sample_no = next((p for p in parts if 'I-N' in p or 'i-n' in p), "")
                            customer = parts[-2] if len(parts) > 2 else ""
                            project = parts[-1] if len(parts) > 1 else ""
                            
                            plan_rows.append({
                                "Seq": seq_val,
                                "Work Order": wo_val,
                                "Acs": parts[2] if len(parts) > 2 else "",
                                "Item Description": item_desc,
                                "Sample Number": sample_no,
                                "Ref.Note": line[100:150].strip() if len(line) > 150 else "",
                                "Pl/Tot.Qty": parts[12] if len(parts) > 12 else "",
                                "Daily Prd.": "",
                                "Yarn Information": line[50:100].strip() if len(line) > 100 else "",
                                "Customer Name": customer,
                                "Project Name": project
                            })
                            
                if plan_rows:
                    df_pdf_plan = pd.DataFrame(plan_rows)
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_plan.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم استخراج وتعبئة بيانات البلان داخل الأعمدة بنجاح ({len(df_pdf_plan)} سجل)!")
                else:
                    st.warning("⚠️ لم يتم العثور على أوردرات مطابقة للنمط داخل ملف الـ PDF.")
        
        # عرض التقرير النهائي الشامل
        st.subheader("📊 التقرير النهائي الشامل بعد توزيع بيانات البلان:")
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