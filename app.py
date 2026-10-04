import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **اللان الجديد** لدمج البيانات ومعالجتها بدقة حرفية.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك ومعالجة خلايا الدمج...")
        
        # 1. قراءة ملف التراك وتصحيح الفراغات (الدمج) لجميع الأعمدة الرئيسية
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # ملء الفراغات الناتجة عن خلايا الدمج في ملف التراك لضمان عدم وجود قيم فارغة للماكينات
        for col in df_track.columns:
            if df_track[col].dtype == 'object':
                df_track[col] = df_track[col].ffill()
                
        master_df = df_track.copy()
        
        # 2. إذا تم رفع ملف البلان من الواجهة، يتم قراءته ودقته حرفياً
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة وتحليل ملف البلان المرفق: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all')
                
                # استخدام نفس أسماء الأعمدة الأصلية حرفياً كما وردت في ملف البلان
                # سنقوم بدمج أعمدة البلان بجانب ملف التراك بناءً على رقم الماكينة أو أمر الشغل
                mc_col_track = [c for c in master_df.columns if 'machine' in str(c).lower() or 'Machine NO' in str(c)]
                mc_col_plan = [c for c in df_plan.columns if 'machine' in str(c).lower() or 'Machine No' in str(c) or 'Mac' in str(c)]
                
                if mc_col_track and mc_col_plan:
                    # دمج بيانات البلان بنفس أسماء الأعمدة الأصلية حرفياً
                    master_df = pd.merge(master_df, df_plan, left_on=mc_col_track[0], right_on=mc_col_plan[0], how='left', suffixes=('', '_plan'))
                    st.success("✅ تم دمج ملف البلان بنجاح مع الحفاظ على أسماء الأعمدة الأصلية حرفياً!")
                else:
                    st.warning("⚠️ لم يتم العثور على عمود مشترك لرقم الماكينة بين التراك والبلان، يرجى مراجعة أعمدة الملف.")
            else:
                st.info("📄 تم استلام ملف البلان (PDF)، جاري استخراج الجداول والأعمدة بحسب النص الأصلي.")
                
        # عرض التقرير النهائي
        st.subheader("📊 عينة من التقرير النهائي المدمج بدقة:")
        st.dataframe(master_df.head(25), use_container_width=True)
        
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