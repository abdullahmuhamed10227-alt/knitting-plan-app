import pandas as pd
import streamlit as st

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("قم برفع **ملف التراك الأساسي (Excel)** وملف **البلان الجديد** لدمج البيانات بدقة تامة ومطابقة أصلية للملفات دون أي تعديل عشوائي.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأصلي...")
        
        # 1. قراءة شيت التتبع الأساسي (OVER VIEW) كما هو من الملف المرفق دون أي تعديل عشوائي
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        master_df = df_track.copy()
        
        # 2. قراءة ملف البلان المرفق حصرياً (إذا تم رفعه) وجلب البيانات والأعمدة الأصلية منه حرفياً
        if uploaded_plan is not None:
            st.info(f"📁 جاري تحليل ملف البلان المرفق: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all')
                
                # البحث عن عمود الماكينة لمطابقة البيانات بدقة حرفية من الملفات المرفقة فقط
                mc_col_track = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'no' in str(c).lower()), None)
                mc_col_plan = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mac' in str(c).lower()), None)
                
                if mc_col_track and mc_col_plan:
                    # دمج بيانات البلان الفعلي بجانب ملف التراك بناءً على المطابقة الفعلية للملفات المرفقة
                    master_df = pd.merge(master_df, df_plan, left_on=mc_col_track, right_on=mc_col_plan, how='left', suffixes=('', '_plan'))
                    st.success("✅ تم دمج بيانات ملف البلان المرفق بدقة حرفية تامة!")
                else:
                    # إذا لم يوجد عمود مشترك، يتم إرفاق أعمدة البلان الأصلي بجانب الجدول دون عشوائية
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
                    st.success("✅ تم إرفاق أعمدة ملف البلان الأصلي بجانب جدول التراك!")
            else:
                st.warning("⚠️ يرجى رفع ملف البلان بصيغة Excel لضمان مطابقة الأعمدة ودمجها بدقة تامة.")
        
        # عرض التقرير النهائي المستخرج من الملفات المرفقة فقط
        st.subheader("📊 التقرير النهائي المستخرج من الملفات المرفقة:")
        st.dataframe(master_df.head(25), use_container_width=True)
        
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
        st.error(f"❌ حدث خطأ أثناء قراءة الملفات: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك (Tracking Excel) وملف البلان للبدء.")