import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك (Excel)** لمعالجة البيانات، ملء الفراغات بدقة، وتوليد التقرير النهائي الشامل.")

uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ومعالجة الملف وتصحيح الفراغات...")
        
        # قراءة الملف الخام لتجاوز مشكلة الخلايا الفارغة في الأوردرات المدمجة
        df_overview = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_overview = df_overview.dropna(how='all')
        
        # ملء الخلايا الفارغة في أعمدة الأوردرات والنوع تلقائياً لتجنب ظهور القيم الفارغة غير المبررة
        work_order_cols = [col for col in df_overview.columns if 'Work ORDER' in str(col) or 'type' in str(col)]
        for col in work_order_cols:
            df_overview[col] = df_overview[col].ffill()
            
        # قراءة شيت أوردرات التسلسل (F.K.G)
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
            
        st.success("✅ تمت معالجة البيانات وتصحيح الفراغات وتوليد التقرير النهائي بنجاح!")
        
        # عرض عينة من الجدول النهائي
        st.subheader("📊 عينة من التقرير النهائي المدمج والمصحح:")
        st.dataframe(master_df.head(25), use_container_width=True)
        
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
        st.error(f"❌ حدث خطأ أثناء معالجة الملفات: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) للبدء في المعالجة.")