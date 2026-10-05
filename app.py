import pandas as pd
import streamlit as st
import pypdf

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الذكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو - التقرير الشامل الموحد")
st.markdown("عرض ملف التراك الأساسي بجوار ملف البلان كاملاً بكل أعمدته ومعلوماته دون أي نقصان أو تعديل.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك الأساسي بكامل أعمدته وهيكله
        df_track_raw = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=None)
        header_track_idx = 1
        for idx, row in df_track_raw.iterrows():
            row_str = str(row.values).lower()
            if 'work order' in row_str or 'machine' in row_str or 'type qualities' in row_str:
                header_track_idx = idx
                break
                
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=header_track_idx)
        df_track = df_track.dropna(how='all')
        
        # تنظيف الدمج للأعمدة الرئيسية في التراك لضمان وضوح الموقف الفعلي
        for col in df_track.columns:
            col_str = str(col).strip().upper()
            if any(k in col_str for k in ['ROLLS', 'ORDER', 'CUSTMER', 'QUALITIES']):
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        
        # 2. قراءة ملف البلان بكامل محتواه دون حذف أي عمود أو معلومة
        df_plan = pd.DataFrame()
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة ملف البلان بكامل تفاصيله: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan_raw = pd.read_excel(uploaded_plan, header=None)
                header_plan_idx = 2
                for idx, row in df_plan_raw.iterrows():
                    row_str = str(row.values).lower()
                    if 'work order' in row_str or 'seq' in row_str:
                        header_plan_idx = idx
                        break
                
                # قراءة ملف البلان كما هو تماماً (كامل الأعمدة)
                df_plan = pd.read_excel(uploaded_plan, header=header_plan_idx)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                # معالجة خلايا الماكينات المدمجة في البلان لملء الفراغات عمودياً لتظهر بجانب كل سطر
                mc_col = None
                for c in df_plan.columns:
                    c_low = str(c).strip().lower()
                    if 'unnamed: 1' in c_low or 'machine' in c_low or 'mc' in c_low or c_low == '1':
                        mc_col = c
                        break
                if not mc_col and len(df_plan.columns) > 1:
                    mc_col = df_plan.columns[1]
                    
                if mc_col:
                    df_plan[mc_col] = df_plan[mc_col].replace('', pd.NA).ffill()
                
                st.success(f"✅ تم تحميل ملف البلان كاملاً ({len(df_plan)} صف) وإضافته بجوار التراك!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                # قراءة نصية دقيقة للـ PDF
                from io import StringIO
                df_plan = pd.read_fwf(StringIO(full_text)) # احتياطي لو ملف نصي أو استخراج سطور
                df_plan = df_plan.fillna("")

        # 3. إلحاق ملف البلان بجوار ملف التراك (دمج أفقي ذكي Full Outer Join يضمن عدم ضياع أي صف من الملفين)
        if not df_plan.empty:
            # توحيد أعمدة المفاتيح للربط (رقم الأوردر والماكينة والسيكوانس)
            track_wo = next((c for c in df_track.columns if 'work order' in str(c).lower() or 'order' in str(c).lower()), None)
            plan_wo = next((c for c in df_plan.columns if 'work order' in str(c).lower() or 'order' in str(c).lower()), None)
            
            track_mc = next((c for c in df_track.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower()), None)
            plan_mc = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower() or 'unnamed: 1' in str(c).lower()), None)

            if track_wo and plan_wo:
                df_track['Key_WO'] = df_track[track_wo].astype(str).str.strip().str.upper()
                df_plan['Key_WO'] = df_plan[plan_wo].astype(str).str.strip().str.upper()
                
                merge_keys = ['Key_WO']
                if track_mc and plan_mc:
                    df_track['Key_MC'] = df_track[track_mc].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                    df_plan['Key_MC'] = df_plan[plan_mc].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                    merge_keys.append('Key_MC')

                # استخدام Full Outer Join لضمان ظهور كل بيانات التراك وكل بيانات البلان بالكامل (بدون حذف أي شيء)
                master_df = pd.merge(df_track, df_plan, on=merge_keys, how='outer', suffixes=('_Tracking', '_Plan'))
                
                for k in ['Key_WO', 'Key_MC']:
                    if k in master_df.columns:
                        master_df = master_df.drop(columns=[k])
            else:
                # لو تعذر الربط، يتم وضع الملفين بجانب بعضهما تسلسلياً لضمان عدم ضياع البيانات
                master_df = pd.concat([df_track.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
        else:
            master_df = df_track

        master_df = master_df.fillna("")

        st.subheader("📊 معاينة التقرير الشامل (التراك + البلان بكامل التفاصيل):")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        output_filename = "Master_Full_Knitting_Report.xlsx"
        master_df.to_excel(output_filename, index=False)
        
        with open(output_filename, "rb") as file:
            st.download_button(
                label="📥 تحميل التقرير الشامل النهائي (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك الأساسي وملف البلان للبدء.")