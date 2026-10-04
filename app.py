import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو - معالجة البيانات واستخراج الأعمدة في مكانها الصحيح بدقة تامة.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك بالكامل
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # معالجة الخلايا المدمجة للأعمدة الإجمالية في التراك
        merge_columns_to_fill = ['ERP ROLLS', 'Roll / Mc  ERP', 'Unnamed: 21', 'Work ORDER', 'type Qualities', 'CUSTMER']
        for col in merge_columns_to_fill:
            if col in df_track.columns:
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة ملف البلان وربطه بذكاء حسب رقم الأوردر (Work Order)
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة بيانات ملف البلان: {uploaded_plan.name}...")
            df_plan = pd.DataFrame()
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                required_columns = ['Work Order', 'Seq', 'Acs', 'Item Description', 'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Daily Prd.', 'Yarn Information', 'Customer Name', 'Project Name']
                for col in required_columns:
                    matching_col = next((c for c in df_plan.columns if col.lower().strip() in str(c).lower().strip()), None)
                    if matching_col and matching_col != col:
                        df_plan = df_plan.rename(columns={matching_col: col})
                
                st.success("✅ تمت قراءة أعمدة وخانات البلان (Excel) بنجاح!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                lines = full_text.split('\n')
                plan_rows = []
                
                for line in lines:
                    clean_line = " ".join(line.split())
                    wo_match = re.search(r'\b\d{6}-\d\b', clean_line)
                    if wo_match:
                        wo_val = wo_match.group()
                        
                        # 1. استخراج السيكوانع (Seq) - أول رقم في السطر غالباً
                        seq_match = re.search(r'^\s*(\d+)', clean_line)
                        seq_val = seq_match.group(1) if seq_match else ""
                        
                        # 2. استخراج رقم العينة (Sample Number)
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+|Sample[A-Za-z0-9\-]+)', clean_line)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        # 3. استخراج اسم المشروع (Project Name) مثل BU1_ORM, BU5_ORM
                        project_match = re.search(r'\b(BU\d+_[A-Za-z0-9]+)\b', clean_line)
                        project_val = project_match.group(1) if project_match else ""
                        
                        # 4. استخراج الكميات والأرقام الأخرى بدقة إذا توفرت في السطر
                        numbers_in_line = re.findall(r'\b\d+\b', clean_line)
                        # استخراج الكمية والإنتاج اليومي بناءً على الأرقام الموجودة بعد السيكوانس
                        qty_val = numbers_in_line[1] if len(numbers_in_line) > 2 else ""
                        daily_val = numbers_in_line[2] if len(numbers_in_line) > 3 else ""

                        words = clean_line.split()
                        # العميل غالباً يكون قبل اسم المشروع أو في نهاية السطر بشكل منظم
                        customer_val = ""
                        for w in words:
                            if w in ["Nike", "Adidas", "Puma", "Zara", "Tommy", "Lacoste", "Under"] or "_" not in w and len(w) > 3 and w != seq_val:
                                customer_val = w

                        plan_rows.append({
                            "Work Order": wo_val,
                            "Seq": seq_val,
                            "Acs": "M" if " M " in clean_line else "",
                            "Item Description": clean_line,
                            "Sample Number": sample_no,
                            "Ref.Note": "",
                            "Pl/Tot.Qty": qty_val,
                            "Daily Prd.": daily_val,
                            "Yarn Information": clean_line[40:120].strip() if len(clean_line) > 120 else "",
                            "Customer Name": customer_val,
                            "Project Name": project_val
                        })
                
                if plan_rows:
                    df_plan = pd.DataFrame(plan_rows)
                    st.success(f"✅ تم استخراج وقراءة محتوى وخانات الـ PDF في مكانها الصحيح ({len(df_plan)} سجل)!")
                else:
                    st.warning("⚠️️ لم يتم العثور على أوردرات مطابقة للنمط داخل ملف الـ PDF.")

            # 3. خطوة الربط الذكي لتفادي الفراغات
            if not df_plan.empty:
                track_wo_col = next((c for c in master_df.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower() or 'order' in str(c).lower()), None)
                plan_wo_col = next((c for c in df_plan.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower()), None)
                
                if track_wo_col and plan_wo_col:
                    master_df[track_wo_col] = master_df[track_wo_col].astype(str).str.strip()
                    df_plan[plan_wo_col] = df_plan[plan_wo_col].astype(str).str.strip()
                    
                    master_df = pd.merge(master_df, df_plan, left_on=track_wo_col, right_on=plan_wo_col, how='left', suffixes=('', '_plan'))
                    st.success("🔗 تم دمج ملف البلان مع التراك بناءً على أرقام الأوردرات بدقة تامة!")
                else:
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)

        master_df = master_df.fillna("")

        st.subheader("📊 التقرير النهائي الشامل (بعد ضبط الأعمدة والبيانات):")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        output_filename = "Master_Knitting_Report_Fixed.xlsx"
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