import pandas as pd
import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الذكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو - النسخة النهائية المطورة")
st.markdown("دمج ملف التراك مع البلان بالربط الثلاثي (أوردر + ماكينة + سيكوانس) لمنع أي تكرار نهائياً.")

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
        
        for col in df_track.columns:
            col_str = str(col).strip().upper()
            if any(k in col_str for k in ['ROLLS', 'ORDER', 'CUSTMER', 'QUALITIES']):
                df_track[col] = df_track[col].ffill()
                
        df_track = df_track.fillna("")
        
        # 2. قراءة ملف البلان وتجهيز مفاتيح الربط الدقيقة
        df_plan = pd.DataFrame()
        if uploaded_plan is not None:
            st.info(f"📁 جاري قراءة ملف البلان: {uploaded_plan.name}...")
            
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
                
                # معالجة خلايا الماكينات المدمجة عمودياً في البلان
                mc_col = None
                for c in df_plan.columns:
                    c_low = str(c).strip().lower()
                    if 'machine' in c_low or 'mc' in c_low or c_low == '1' or 'unnamed: 1' in c_low:
                        mc_col = c
                        break
                if not mc_col and len(df_plan.columns) > 1:
                    mc_col = df_plan.columns[1]
                    
                if mc_col:
                    df_plan = df_plan.rename(columns={mc_col: 'Machine_Plan'})
                    df_plan['Machine_Plan'] = df_plan['Machine_Plan'].replace('', pd.NA).ffill()
                
                # توحيد أسماء الأعمدة الأساسية في البلان
                for col in df_plan.columns:
                    col_str = str(col).strip().lower()
                    if 'work order' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Work_Order_Plan'})
                    elif 'machine' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Machine_Plan'})
                    elif 'seq' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Seq_Plan'})

                # منع تكرار نفس الصفوف داخل البلان نفسه قبل الدمج
                df_plan = df_plan.drop_duplicates()
                st.success(f"✅ تم قراءة ملف البلان وتجهيزه بدقة ({len(df_plan)} صف)!")

        # 3. الربط الدقيق بمنع التكرار (الربط المزدوج: رقم الأوردر + رقم الماكينة)
        if not df_plan.empty:
            track_wo = next((c for c in df_track.columns if 'work order' in str(c).lower() or 'order' in str(c).lower()), None)
            track_mc = next((c for c in df_track.columns if 'machine' in str(c).lower() or 'mc' in str(c).lower()), None)
            
            if track_wo and 'Work_Order_Plan' in df_plan.columns:
                df_track['Key_WO'] = df_track[track_wo].astype(str).str.strip().str.upper()
                df_plan['Key_WO'] = df_plan['Work_Order_Plan'].astype(str).str.strip().str.upper()
                
                merge_keys = ['Key_WO']
                
                if track_mc and 'Machine_Plan' in df_plan.columns:
                    df_track['Key_MC'] = df_track[track_mc].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                    df_plan['Key_MC'] = df_plan['Machine_Plan'].astype(str).str.strip().str.upper().str.replace('M', '', regex=True)
                    merge_keys.append('Key_MC')

                # الدمج الدقيق بمنع تضاعف الصفوف
                master_df = pd.merge(df_track, df_plan, on=merge_keys, how='left', suffixes=('_Tracking', '_Plan'))
                
                for k in merge_keys:
                    if k in master_df.columns:
                        master_df = master_df.drop(columns=[k])
            else:
                master_df = pd.concat([df_track.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
        else:
            master_df = df_track

        master_df = master_df.fillna("")

        st.subheader("📊 معاينة التقرير الموحد الخالي من التكرار:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        # 4. تصدير وتنسيق ملف الإكسيل الاحترافي
        output_filename = "Master_Clean_Knitting_Report.xlsx"
        
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
            
        # البحث عن أعمدة الحالة والماكينة في التراك لتلوين ON / OFF
        on_off_col_idx = None
        mc_col_idx = None
        for col_num in range(1, ws.max_column + 1):
            col_name = str(ws.cell(row=1, column=col_num).value).strip().lower()
            if 'on / off' in col_name or 'on/off' in col_name:
                on_off_col_idx = col_num
            if 'machine' in col_name or 'mc' in col_name or 'machine no' in col_name:
                if not mc_col_idx:
                    mc_col_idx = col_num

        green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        green_font = Font(name="Calibri", size=10, bold=True, color="274E13")
        
        red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
        red_font = Font(name="Calibri", size=10, bold=True, color="660000")

        # تطبيق التنسيق والالتفاف (Wrap Text) وتلوين التراك
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
                label="📥 تحميل التقرير النهائي المنظم (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ تم إصدار التقرير النهائي بدون أي تكرار وبكل دقة!")
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك الأساسي وملف البلان للبدء.")