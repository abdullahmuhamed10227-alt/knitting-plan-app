import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك (Excel)** وملف **البلان الجديد** يومياً لمعالجة البيانات وتوليد التقرير النهائي.")

# --- قسم رفع الملفات من الواجهة ---
col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (PDF أو Excel)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ومعالجة الملفات...")
        
        # قراءة شيت التتبع الأساسي (OVER VIEW) من ملف التراك المرفق
        df_overview = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_overview = df_overview.dropna(how='all')
        
        # قراءة شيت أوردرات التسلسل (F.K.G) من ملف التراك
        try:
            df_fkg = pd.read_excel(uploaded_tracking, sheet_name='F.K.G')
        except Exception:
            df_fkg = pd.DataFrame()
            
        # معالجة وتوزيع أوردرات البلان في أعمدة جديدة على اليمين (حسب السيكونس Sira)
        if not df_fkg.empty and 'Machine No' in df_fkg.columns and 'Sira' in df_fkg.columns:
            df_fkg_pivot = df_fkg.pivot_table(
                index='Machine No',
                columns='Sira',
                values=['Work Order Name', 'is Emri Kalemi', 'Fabric Code', 'Planned Qty', 'Remained Qty', 'Release Date'],
                aggfunc='first'
            )
            df_fkg_pivot.columns = [f"Seq_{s}_{val}" for val, s in df_fkg_pivot.columns]
            df_fkg_pivot = df_fkg_pivot.reset_index()
            
            mc_col_candidates = [c for c in df_overview.columns if 'machine' in str(c).lower() or 'Machine NO' in str(c)]
            if mc_col_candidates:
                mc_col = mc_col_candidates[0]
                master_df = pd.merge(df_overview, df_fkg_pivot, left_on=mc_col, right_on='Machine No', how='left')
            else:
                master_df = df_overview
        else:
            master_df = df_overview
            
        st.success("✅ تمت معالجة البيانات وتوليد التقرير النهائي بنجاح!")
        
        # عرض عينة من الجدول النهائي
        st.subheader("📊 عينة من التقرير النهائي المدمج:")
        st.dataframe(master_df.head(20))
        
        # حفظ الملف وتوفير زر التحميل
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
        st.error(f"❌ حدث خطأ أثناء قراءة أو معالجة الملفات: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) على الأقل للبدء في العرض والمعالجة.")