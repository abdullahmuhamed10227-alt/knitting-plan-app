import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام إدارة ومتابعة الـ 156 ماكينة ودمج أوردرات البلان بدقة تامة.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي لكافة الماكينات...")
        
        # 1. قراءة ملف التراك بالكامل (الـ 156 ماكينة)
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # معالجة الخلايا المدمجة للأعمدة الإجمالية والأوردرات
        merge_columns_to_fill = ['ERP ROLLS', 'Roll / Mc  ERP', 'Unnamed: 21', 'Work ORDER', 'type Qualities', 'CUSTMER']
        for col in merge_columns_to_fill:
            if col in df_track.columns:
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        total_machines = len(master_df)
        st.success(f"✅ تم تحميل ملف التراك بنجاح! إجمالي الماكينات المسجلة في التقرير: **{total_machines} ماكينة**")
        
        # البحث عن عمود رقم الماكينة في التراك
        mc_col_track = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'no' in str(c).lower()), None)
        
        # 2. قراءة ملف البلان المرفق ومطابقته برقم الماكينة
        if uploaded_plan is not None and mc_col_track:
            st.info(f"📁 جاري تحليل ملف البلان المرفق: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                mc_col_plan = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mac' in str(c).lower() or 'no' in str(c).lower()), None)
                
                if mc_col_plan:
                    master_df = pd.merge(master_df, df_plan, left_on=mc_col_track, right_on=mc_col_plan, how='left', suffixes=('', '_plan'))
                    st.success("✅ تم ربط ودمج أوردرات البلان لكل ماكينة بدقة تامة!")
                else:
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                    st.success("✅ تمت إضافة أعمدة البلان بجانب جدول الماكينات!")
                    
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                plan_rows = []
                for page_idx, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    for line in text.split('\n'):
                        wo_match = re.search(r'\b\d{6}-\d\b', line)
                        if wo_match:
                            plan_rows.append({"Work Order (Plan)": wo_match.group(), "PDF Page": page_idx + 1, "Details": line})
                            
                if plan_rows:
                    df_pdf_plan = pd.DataFrame(plan_rows)
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_plan.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم استخراج أوردرات البلان من الـ PDF وإضافتها بنجاح ({len(df_pdf_plan)} أوردر)!")
        
        # عرض التقرير النهائي الشامل لكافة الـ 156 ماكينة
        st.subheader("📊 التقرير النهائي الشامل لجميع الماكينات:")
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