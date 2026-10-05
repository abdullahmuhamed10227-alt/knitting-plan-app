import pandas as pd
import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الذكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو - التقرير المستقل والمنظم")
st.markdown("عرض ملف التراك منسقاً وملوناً، وبجواره ملف البلان مرتباً بالترتيب النموذجي الدقيق بدون أي تداخل أو تكرار في الأسماء.")

col1, col2 = st.columns(2)

with col1:
    uploaded_tracking = st.file_uploader("📂 ارفع ملف التراك الأساسي (Excel)", type=["xlsx"])

with col2:
    uploaded_plan = st.file_uploader("📂 ارفع ملف البلان الجديد (Excel أو PDF)", type=["xlsx", "pdf"])

if uploaded_tracking is not None and uploaded_plan is not None:
    try:
        st.info("🔄 جاري معالجة ملفات التراك والبلان باستقلالية تامة...")
        
        # 1. قراءة ومعالجة ملف التراك الأساسي
        df_track_raw = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=None)
        header_track_idx = 1
        for idx, row in df_track_raw.iterrows():
            row_str = str(row.values).lower()
            if 'work order' in row_str or 'machine' in row_str or 'type qualities' in row_str:
                header_track_idx = idx
                break
                
        df_track = pd.read_excel(uploaded_tracking, sheet_name='OVER VIEW', header=header_track_idx)
        df_track = df_track.dropna(how='all')
        
        for col in df_track.columns:
            col_str = str(col).strip().upper()
            if any(k in col_str for k in ['ROLLS', 'ORDER', 'CUSTMER', 'QUALITIES']):
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        
        # تنظيف أسماء أعمدة التراك وجعلها فريدة تماماً لمنع أي تكرار
        track_columns = []
        seen_track = set()
        for idx, c in enumerate(df_track.columns):
            c_name = str(c).strip()
            if not c_name or c_name.lower().startswith('unnamed'):
                c_name = f"Track_Col_{idx}"
            base_name = f"{c_name}_Tracking"
            while base_name in seen_track:
                base_name += f"_{idx}"
            seen_track.add(base_name)
            track_columns.append(base_name)
        df_track.columns = track_columns
        
        # 2. قراءة ومعالجة ملف البلان وترتيب أعمدته بالترتيب النموذجي المطلوب
        df_plan = pd.DataFrame()
        if uploaded_plan.name.endswith('.xlsx'):
            df_plan_raw = pd.read_excel(uploaded_plan, header=None)
            header_plan_idx = 0
            for idx, row in df_plan_raw.iterrows():
                row_str = str(row.values).lower()
                if 'work order' in row_str or 'machine' in row_str:
                    header_plan_idx = idx
                    break
            
            df_plan = pd.read_excel(uploaded_plan, header=header_plan_idx)
            df_plan = df_plan.dropna(how='all').fillna("")
            
            rename_map = {}
            for col in df_plan.columns:
                c_str = str(col).strip().lower()
                if 'work order' in c_str:
                    rename_map[col] = 'Work Order'
                elif 'machine' in c_str or c_str == '1' or 'unnamed: 1' in c_str:
                    rename_map[col] = 'Machine'
                elif 'seq' in c_str:
                    rename_map[col] = 'Seq.'
                elif 'gauge' in c_str or 'specs' in c_str:
                    rename_map[col] = 'Gauge_Specs'
                elif 'item' in c_str:
                    rename_map[col] = 'Item Description'
                elif 'sample' in c_str:
                    rename_map[col] = 'Sample Number'
                elif 'ref' in c_str:
                    rename_map[col] = 'Ref.Note'
                elif 'qty' in c_str or 'tot' in c_str:
                    rename_map[col] = 'Pl/Tot.Qty'
                elif 'daily' in c_str or 'prd' in c_str:
                    rename_map[col] = 'Dailiy Prd.'
                elif 'yarn' in c_str:
                    rename_map[col] = 'Yarn Information'
                elif 'customer' in c_str:
                    rename_map[col] = 'Customer Name'
                elif 'project' in c_str:
                    rename_map[col] = 'Project Name'
            
            df_plan = df_plan.rename(columns=rename_map)
            
            standard_plan_cols = [
                'Work Order', 'Machine', 'Seq.', 'Gauge_Specs', 'Item Description', 
                'Sample Number', 'Ref.Note', 'Pl/Tot.Qty', 'Dailiy Prd.', 
                'Yarn Information', 'Customer Name', 'Project Name'
            ]
            
            for std_col in standard_plan_cols:
                if std_col not in df_plan.columns:
                    df_plan[std_col] = ""
                    
            df_plan = df_plan[standard_plan_cols]
            
            if 'Machine' in df_plan.columns:
                df_plan['Machine'] = df_plan['Machine'].replace('', pd.NA).ffill()
                
            df_plan = df_plan.drop_duplicates()

        df_plan = df_plan.fillna("")
        
        # تنظيف أسماء أعمدة البلان وجعلها فريدة تماماً لمنع أي تكرار
        plan_columns = []
        seen_plan = set()
        for idx, c in enumerate(df_plan.columns):
            c_name = str(c).strip()
            base_name = f"{c_name}_Plan"
            while base_name in seen_plan:
                base_name += f"_{idx}"
            seen_plan.add(base_name)
            plan_columns.append(base_name)
        df_plan.columns = plan_columns

        # 3. عرض الجدولين جنباً إلى جنب ككتلتين مستقلتين تماماً
        df_track_reset = df_track.reset_index(drop=True)
        df_plan_reset = df_plan.reset_index(drop=True)
        
        df_track_reset['--- TRACKING (الموقف الفعلي) ---'] = ""
        df_plan_reset['--- PLAN (خطة التشغيل) ---'] = ""
        
        master_df = pd.concat([df_track_reset, df_plan_reset], axis=1)
        master_df = master_df.fillna("")

        st.subheader("📊 معاينة التقرير المستقل بالترتيب القياسي:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        # 4. تصدير وتنسيق ملف الإكسيل الاحترافي
        output_filename = "Master_Standard_Knitting_Report.xlsx"
        
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            master_df.to_excel(writer, index=False, sheet_name='Master Report')
            
        wb = openpyxl.load_workbook(output_filename)
        ws = wb.active
        
        # تنسيق الهيدر
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        # تلوين حالات ON (أخضر) و OFF (أحمر) في جزء التراك فقط
        on_off_col_idx = None
        mc_col_idx = None
        for col_num in range(1, ws.max_column + 1):
            col_name = str(ws.cell(row=1, column=col_num).value).strip().lower()
            if 'on / off' in col_name or 'on/off' in col_name:
                on_off_col_idx = col_num
            if ('machine' in col_name or 'mc' in col_name) and '_tracking' in col_name:
                if not mc_col_idx:
                    mc_col_idx = col_num

        green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        green_font = Font(name="Calibri", size=10, bold=True, color="274E13")
        
        red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
        red_font = Font(name="Calibri", size=10, bold=True, color="660000")

        for row_num in range(2, ws.max_row + 1):
            status_val = ""
            if on_off_col_idx:
                status_val = str(ws.cell(row=row_num, column=on_off_col_idx).value).strip().upper()
                
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_num, column=col_num)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                
                if status_val == "ON":
                    if col_num == on_off_col_idx or (mc_col_idx and col_num == mc_col_idx):
                        cell.fill = green_fill
                        cell.font = green_font
                elif status_val == "OFF":
                    if col_num == on_off_col_idx or (mc_col_idx and col_num == mc_col_idx):
                        cell.fill = red_fill
                        cell.font = red_font

        # ضبط عرض الأعمدة تلقائياً
        for col in ws.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    val_str = str(cell.value)
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        wb.save(output_filename)

        with open(output_filename, "rb") as file:
            st.download_button(
                label="📥 تحميل التقرير القياسي النهائي (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ تم إصدار التقرير بالترتيب القياسي الدقيق وبفصل تام بين التراك والبلان!")
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
else:
    st.info("ℹ️ يرجى رفع ملف التراك وملف البلان للبدء.")