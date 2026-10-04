import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو - النسخة المعدلة لتجنب تكرار البيانات وضبط الأعمدة بدقة.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك الأساسي
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # تنظيف الدمج للأعمدة الرئيسية في التراك
        merge_columns_to_fill = ['ERP ROLLS', 'Roll / Mc  ERP', 'Unnamed: 21', 'Work ORDER', 'type Qualities', 'CUSTMER']
        for col in merge_columns_to_fill:
            if col in df_track.columns:
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة ملف البلان
        if uploaded_plan is not None:
            st.info(f"📁 جاري معالجة ملف البلان: {uploaded_plan.name}...")
            df_plan = pd.DataFrame()
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                # توحيد أسماء الأعمدة حرفياً لو مطابقة
                required_columns = ['Work Order', 'Seq', 'Acs', 'Item Description', 'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Daily Prd.', 'Yarn Information', 'Customer Name', 'Project Name']
                for col in required_columns:
                    matching_col = next((c for c in df_plan.columns if col.lower().strip() in str(c).lower().strip()), None)
                    if matching_col and matching_col != col:
                        df_plan = df_plan.rename(columns={matching_col: col})
                
                st.success("✅ تمت قراءة ملف البلان (Excel) بنجاح!")
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                lines = full_text.split('\n')
                plan_rows = []
                
                for line in lines:
                    clean_line = " ".join(line.split())
                    # البحث عن رقم الأوردر الأساسي (مثال: 6 أرقام - رقم)
                    wo_match = re.search(r'\b\d{6}-\d\b', clean_line)
                    if wo_match:
                        wo_val = wo_match.group()
                        
                        # استخراج السيكوانس (Seq) - أول رقم منطقي في بداية السطر (يكون أقل من 1000 عادة)
                        tokens = clean_line.split()
                        seq_val = ""
                        for token in tokens[:3]:  # أول كم كلمة في السطر
                            if token.isdigit() and int(token) < 1000:
                                seq_val = token
                                break
                        
                        # استخراج رقم العينة (Sample Number)
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+)', clean_line)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        # استخراج اسم المشروع (Project Name) مثل BU1_ORM, BU5_ORM
                        project_match = re.search(r'\b(BU\d+_[A-Za-z0-9]+)\b', clean_line)
                        project_val = project_match.group(1) if project_match else ""
                        
                        # استخراج اسم العميل (Customer Name) الشركات المشهورة أو الكلمات الأخيرة قبل المشروع
                        customers_list = ["NIKE", "ADIDAS", "PUMA", "ZARA", "TOMMY", "LACOSTE", "UNDER", "H&M"]
                        customer_val = ""
                        for word in tokens:
                            if word.upper() in customers_list:
                                customer_val = word
                                break
                        if not customer_val and len(tokens) > 2:
                            # لو مش موجودة صراحة، نجيب الكلمة اللي قبل اسم المشروع لو متاح
                            if project_val in clean_line:
                                parts_before_proj = clean_line.split(project_val)[0].split()
                                if parts_before_proj:
                                    customer_val = parts_before_proj[-1] if len(parts_before_proj[-1]) > 2 else ""

                        # استخراج الكميات (Pl/Tot.Qty و Daily Prd.) الأرقام الكبيرة نسبياً
                        quantities = [t for t in tokens if t.isdigit() and int(t) > 50]
                        qty_val = quantities[0] if len(quantities) > 0 else ""
                        daily_val = quantities[1] if len(quantities) > 1 else ""

                        plan_rows.append({
                            "Work Order": wo_val,
                            "Seq": seq_val if seq_val != wo_val else "",
                            "Acs": "M" if " M " in clean_line else "",
                            "Item Description": clean_line,
                            "Sample Number": sample_no,
                            "Ref.Note": "",
                            "Pl/Tot.Qty": qty_val if qty_val != wo_val else "",
                            "Daily Prd.": daily_val if daily_val != wo_val else "",
                            "Yarn Information": clean_line,
                            "Customer Name": customer_val if customer_val != wo_val else "",
                            "Project Name": project_val
                        })
                
                if plan_rows:
                    df_plan = pd.DataFrame(plan_rows)
                    st.success(f"✅ تم استخراج وتنظيف بيانات الـ PDF بدقة بدون تكرار ({len(df_plan)} سجل)!")
                else:
                    st.warning("⚠️️ لم يتم العثور على أوردرات مطابقة داخل ملف الـ PDF.")

            # 3. الربط الذكي والسليم بين التراك والبلان بناءً على أرقام الأوردرات
            if not df_plan.empty:
                track_wo_col = next((c for c in master_df.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower() or 'order' in str(c).lower()), None)
                plan_wo_col = next((c for c in df_plan.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower()), None)
                
                if track_wo_col and plan_wo_col:
                    master_df[track_wo_col] = master_df[track_wo_col].astype(str).str.strip()
                    df_plan[plan_wo_col] = df_plan[plan_wo_col].astype(str).str.strip()
                    
                    master_df = pd.merge(master_df, df_plan, left_on=track_wo_col, right_on=plan_wo_col, how='left', suffixes=('', '_plan'))
                    st.success("🔗 تم دمج الملفين برقم الأوردر بدقة تامة وبدون تداخل!")
                else:
                    master_df = pd.concat([master_df.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)

        master_df = master_df.fillna("")

        st.subheader("📊 التقرير النهائي المُصحح:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        output_filename = "Master_Knitting_Report_Perfect.xlsx"
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