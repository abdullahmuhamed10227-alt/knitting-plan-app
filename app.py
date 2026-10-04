import pandas as pd
import streamlit as st
import pypdf
import re

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الديناميكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو الديناميكي")
st.markdown("نظام تخطيط ومتابعة التريكو - النسخة النهائية المظبوطة للربط الدقيق بالماكينة ورقم الأوردر.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None:
    try:
        st.info("🔄 جاري قراءة ملف التراك الأساسي...")
        
        # 1. قراءة ملف التراك الأساسي والحفاظ على هيكله
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=1)
        df_track = df_track.dropna(how='all')
        
        # تنظيف الدمج للأعمدة الرئيسية في التراك
        merge_columns_to_fill = ['ERP ROLLS', 'Roll / Mc  ERP', 'Unnamed: 21', 'Work ORDER', 'type Qualities', 'CUSTMER']
        for col in merge_columns_to_fill:
            if col in df_track.columns:
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        master_df = df_track.copy()
        
        # 2. قراءة ملف البلان بمعزل تام
        df_plan = pd.DataFrame()
        if uploaded_plan is not None:
            st.info(f"📁 جاري معالجة ملف البلان: {uploaded_plan.name}...")
            
            if uploaded_plan.name.endswith('.xlsx'):
                df_plan = pd.read_excel(uploaded_plan)
                df_plan = df_plan.dropna(how='all').fillna("")
                
                required_columns = ['Work Order', 'Machine', 'Seq', 'Acs', 'Item Description', 'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Daily Prd.', 'Yarn Information', 'Customer Name', 'Project Name']
                for col in required_columns:
                    matching_col = next((c for c in df_plan.columns if col.lower().strip() in str(c).lower().strip()), None)
                    if matching_col and matching_col != col:
                        df_plan = df_plan.rename(columns={matching_col: col})
                
            elif uploaded_plan.name.endswith('.pdf'):
                pdf_reader = pypdf.PdfReader(uploaded_plan)
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text() + "\n"
                
                lines = full_text.split('\n')
                plan_rows = []
                
                for line in lines:
                    clean_line = " ".join(line.split())
                    if not clean_line:
                        continue
                        
                    wo_match = re.search(r'\b\d{6}-\d\b', clean_line)
                    machine_match = re.search(r'\b(M\d+|\d{3,4})\b', clean_line)
                    
                    if wo_match:
                        wo_val = wo_match.group()
                        machine_val = machine_match.group() if machine_match else ""
                        
                        tokens = clean_line.split()
                        seq_val = ""
                        for token in tokens[:3]:
                            if token.isdigit() and int(token) < 1000:
                                seq_val = token
                                break
                        
                        sample_match = re.search(r'(I-N[A-Za-z0-9\-]+|i-n[A-Za-z0-9\-]+)', clean_line)
                        sample_no = sample_match.group(1) if sample_match else ""
                        
                        project_match = re.search(r'\b(BU\d+_[A-Za-z0-9]+)\b', clean_line)
                        project_val = project_match.group(1) if project_match else ""
                        
                        customers_list = ["NIKE", "ADIDAS", "PUMA", "ZARA", "TOMMY", "LACOSTE", "UNDER", "H&M"]
                        customer_val = ""
                        for word in tokens:
                            if word.upper() in customers_list:
                                customer_val = word
                                break
                        
                        quantities = [t for t in tokens if t.isdigit() and int(t) > 50]
                        qty_val = quantities[0] if len(quantities) > 0 else ""
                        daily_val = quantities[1] if len(quantities) > 1 else ""

                        plan_rows.append({
                            "Work Order": wo_val,
                            "Machine": machine_val,
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

            # 3. الربط الدقيق والمزدوج (بين الماكينة ورقم الأوردر معاً)
            if not df_plan.empty:
                track_wo_col = next((c for c in master_df.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower() or 'order' in str(c).lower()), None)
                track_mc_col = next((c for c in master_df.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower()), None)
                
                plan_wo_col = next((c for c in df_plan.columns if 'work order' in str(c).lower() or 'wo' in str(c).lower()), None)
                plan_mc_col = next((c for c in df_plan.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower()), None)
                
                if track_wo_col and plan_wo_col:
                    # توحيد صيغة مفاتيح الربط لتتطابق تماماً في الطرفين
                    master_df['Clean_WO'] = master_df[track_wo_col].astype(str).str.strip().str.upper()
                    df_plan['Clean_WO'] = df_plan[plan_wo_col].astype(str).str.strip().str.upper()
                    
                    merge_keys = ['Clean_WO']
                    
                    if track_mc_col and plan_mc_col:
                        master_df['Clean_MC'] = master_df[track_mc_col].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                        df_plan['Clean_MC'] = df_plan[plan_mc_col].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                        merge_keys.append('Clean_MC')

                    # إزالة أعمدة البلان القديمة من التراك لو وجدت لمنع التداخل
                    for c in ['Work Order', 'Machine', 'Seq', 'Acs', 'Item Description', 'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Daily Prd.', 'Yarn Information', 'Customer Name', 'Project Name']:
                        if c in master_df.columns:
                            master_df = master_df.drop(columns=[c])

                    master_df = pd.merge(master_df, df_plan, on=merge_keys, how='left')
                    
                    # تنظيف الأعمدة المؤقتة للربط
                    for temp_col in ['Clean_WO', 'Clean_MC']:
                        if temp_col in master_df.columns:
                            master_df = master_df.drop(columns=[temp_col])
                        if temp_col in df_plan.columns:
                            df_plan = df_plan.drop(columns=[temp_col])
                            
                    st.success("✅ تم ربط بيانات البلان بالماكينة ورقم الأوردر بدقة تامة ودون أي تكرار خاطئ!")

        master_df = master_df.fillna("")

        st.subheader("📊 التقرير النهائي المنظم:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        output_filename = "Master_Knitting_Report_Accurate.xlsx"
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