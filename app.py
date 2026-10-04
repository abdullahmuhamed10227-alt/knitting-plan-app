import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **البلان الجديد (Excel أو PDF)** لمعالجة البيانات بدقة تامة.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (PDF أو Excel)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك كما هو من الملف المرفق دون أي تعديل أو عشوائية
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # تنظيف القيم الفارغة لتظهر بشكل نظيف تماماً مثل أصل الإكسيل
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. معالجة ملف البلان المرفق (سواء كان Excel أو PDF بدقة عالية)
        if uploaded_plan is not None:
            st.info(f"📁 جاري تحليل ملف البلان المرفق: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                # دمج بيانات البلان مع التراك بناءً على المطابقة الفعلية
                mc_col_track = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'no' in str(c).lower()), None)
                mc_col_plan = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mac' in str(c).lower()), None)
                
                if mc_col_track and mc_col_plan:
                    master_df = pd.merge(master_df, df_plan, left_on=mc_col_track, right_on=mc_col_plan, how='left', suffixes=('', '_plan'))
                    st.success("✅ تم دمج ملف البلان (Excel) بدقة تامة!")
                else:
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                    st.success("✅ تم إرفاق أعمدة ملف البلان بجانب جدول التراك!")
                    
            elif uploaded_plan.name.endswith('.pdf'):
                # قراءة واستخراج البيانات من ملف الـ PDF بكفاءة عالية
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                extracted_data = []
                
                for page_num, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    lines = text.split('\n')
                    for line in lines:
                        # استخراج أوردرات الشغل من ملف الـ PDF
                        wo_match = re.search(r'\b\d{6}-\d\b', line)
                        if wo_match:
                            extracted_data.append({
                                "PDF_Page": page_num + 1,
                                "Extracted_Work_Order": wo_match.group(),
                                "PDF_Text_Line": line
                            })
                            
                if extracted_data:
                    df_pdf_extracted = pd.DataFrame(extracted_data)
                    # إرفاق الأوردرات المستخرجة من الـ PDF بجانب التقرير النهائي
                    master_df = pd.concat([master_df.reset_index(drop=True), df_pdf_extracted.reset_index(drop=True)], axis=1)
                    st.success(f"✅ تم قراءة واستخراج أوردرات البلان (PDF) بنجاح ({len(df_pdf_extracted)} أوردر مستخرج)!")
                else:
                    st.warning("⚠️ لم يتم العثور على أوردرات تطابق النمط (XXXXXX-X) في ملف الـ PDF المرفق.")
        
        # عرض التقرير النهائي
        st.subheader("📊 التقرير النهائي المستخرج من الملفات المرفقة:")
        st.dataframe(master_df.head(25), use_container_width=True)
        
        # زر التحميل المباشر بصيغة Excel
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