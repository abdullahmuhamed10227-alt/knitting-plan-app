import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **اللان (Excel أو PDF)** لاستخراج الأعمدة المحددة بدقة حرفية.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي (الـ 156 ماكينة)...")
        
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
        
        # 2. قراءة وتحليل ملف البلان واستخراج الأعمدة المطلوبة بدقة
        if uploaded_plan is not None:
            st.info(f"📁 جاري استخراج الأعمدة المحددة من ملف البلان: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                st.success("✅ تمت إضافة أعمدة ملف البلان (Excel) بدقة تامة!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                plan_rows = []
                
                for page_idx, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    lines = text.split('\n')
                    for line in lines:
                        # البحث عن نمط أوردر الشغل لاستخراج البيانات المرتبطة به
                        wo_match = re.search(r'\b\d{6}-\d\b', line)
                        if wo_match:
                            # استخراج السيكونس والأوردر وباقي التفاصيل من السطر
                            parts = line.split()
                            seq_val = parts[0] if parts and parts[0].isdigit() else ""
                            wo_val = wo_match.group()
                            
                            plan_rows.append({
                                "Seq": seq_val,
                                "Work Order": wo_val,
                                "Acs": parts[2] if len(parts) > 2 else "",
                                "Item Description": line,
                                "Sample Number": "",
                                "Ref.Note": "",
                                "Pl/Tot.Qty": "",
                                "Daily Prd.": "",
                                "Yarn Information": "",
                                "Customer Name": "",
                                "Project Name": ""
                            })
                            
                if plan_rows:
                    df_pdf_plan = pd.DataFrame(plan_rows)
                    # إرفاق الأعمدة المحددة بجانب جدول التراك
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_plan.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم استخراج وتوزيع الأعمدة المطلوبة من ملف الـ PDF بنجاح ({len(df_pdf_plan)} سجل)!")
                else:
                    st.warning("⚠️ لم يتم التعرف على الأوردرات بالشكل المطلوب داخل ملف الـ PDF.")
        
        # عرض التقرير النهائي
        st.subheader("📊 التقرير النهائي الشامل بعد دمج أعمدة البلان:")
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