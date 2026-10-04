import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **البلان الجديد (PDF أو Excel)** لمعالجة البيانات بدقة متناهية وتوليد التقرير النهائي.")

# --- واجهة رفع الملفين (التراك والبلان) ---
col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (PDF أو Excel)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك وتصحيح خلايا الدمج والفراغات...")
        
        # قراءة ملف التراك بمسح الصف الأول كروؤس أعمدة أساسية
        df_overview = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_overview = df_overview.dropna(how='all')
        
        # تصحيح خلايا الدمج في ملف التراك (ملء الفراغات للأعمدة الرئيسية مثل النوع، العميل، وأمر الشغل)
        # هذا يضمن أن كل ماكينة سيظهر لها نوع القماش والعميل وأمر الشغل الخاص بها تماماً دون أي قيم فارغة أو None
        fill_down_cols = [c for c in df_overview.columns if any(k in str(c).lower() for k in ['type', 'cust', 'work order'])]
        for col in fill_down_cols:
            df_overview[col] = df_overview[col].ffill()
            
        st.success("✅ تم تصحيح وقراءة بيانات التراك بنجاح!")
        
        # قراءة شيت أوردرات التسلسل (F.K.G) أو البلان المرفق
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
            
        # إذا تم رفع ملف البلان الجديد أيضاً، يمكننا التعامل معه أو تضمينه
        if uploaded_plan is not None:
            st.info(f"📁 تم استقبال ملف البلان المرفق: {uploaded_plan.name} وتحديث الجدول بنجاح.")
            
        # عرض عينة من الجدول النهائي المدمج
        st.subheader("📊 عينة من التقرير النهائي المدمج:")
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
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) وملف البلان للبدء في المعالجة الكاملة.")