import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **البلان الجديد** لمعالجة البيانات ومطابقتها تماماً لأصل الملفات.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك ومعالجة الخلايا بدقة تامة...")
        
        # 1. قراءة ملف التراك مع الحفاظ على الهيكل الأصلي
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # استبدال القيم الفارغة لتعرض بشكل طبيعي مطابق للإكسيل
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة ملف البلان المرفق (إن وجد) ودمجه
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة وتحليل ملف البلان: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                mc_col_track = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'no' in str(c).lower()), None)
                mc_col_plan = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mac' in str(c).lower()), None)
                
                if mc_col_track and mc_col_plan:
                    master_df = pd.merge(master_df, df_plan, left_on=mc_col_track, right_on=mc_col_plan, how='left', suffixes=('', '_plan'))
                    st.success("✅ تم دمج ملف البلان بنجاح تام!")
                else:
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                    st.success("✅ تم إرفاق أعمدة البلان بجوار جدول التراك!")
            else:
                st.warning("⚠️ يرجى رفع ملف البلان بصيغة Excel لضمان مطابقة الأعمدة ودمجها بدقة.")
        
        # عرض الجدول النهائي المطابق
        st.subheader("📊 التقرير النهائي المطابق لأصل الملفات:")
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
        st.error(f"❌ حدث خطأ أثناء قراءة الملفات: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) وملف البلان للبدء.")