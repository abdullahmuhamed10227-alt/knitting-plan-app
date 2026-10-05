import pandas as pd
import streamlit as st
import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="نظام تخطيط ومتابعة التريكو الذكي", layout="wide")
st.title("🧵 نظام تخطيط ومتابعة التريكو - التقرير الموحد المطابق للنموذج المعتمد")
st.markdown("دمج ملف التراك مع ملف البلان بحيث تظهر الأوردرات والماكينات بالهيكل الاحترافي المطلوب.")

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
        
        # 2. قراءة ملف البلان مع الحفاظ على هيكله النموذجي (Work Order مدمج والماكينات تحت بعضها)
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
                
                # توحيد أسماء الأعمدة الأساسية في البلان مطابقة للنموذج
                for col in df_plan.columns:
                    col_str = str(col).strip().lower()
                    if 'work order' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Work Order'})
                    elif 'machine' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Machine'})
                    elif 'seq' in col_str:
                        df_plan = df_plan.rename(columns={col: 'Seq.'})

                st.success(f"✅ تم قراءة ملف البلان بهيكله النموذجي بنجاح ({len(df_plan)} صف)!")

        # 3. الربط بين التراك والبلان مع الحفاظ على هيكل عرض البلان بجوار التراك
        if not df_plan.empty:
            track_wo = next((c for c in df_track.columns if 'work order' in str(c).lower() or 'order' in str(c).lower()), None)
            
            if track_wo and 'Work Order' in df_plan.columns:
                # إنشاء نسخة مؤقتة مملوءة للأوردرات لغرض المطابقة الدقيقة فقط دون الإخلال بعرض الملف الأصلي
                df_plan['Temp_WO_Fill'] = df_plan['Work Order'].replace('', pd.NA).ffill()
                df_track['Temp_WO_Fill'] = df_track[track_wo].astype(str).str.strip().str.upper()
                df_plan['Temp_WO_Fill'] = df_plan['Temp_WO_Fill'].astype(str).str.strip().str.upper()
                
                master_df = pd.merge(df_track, df_plan, on='Temp_WO_Fill', how='outer', suffixes=('_Tracking', '_Plan'))
                if 'Temp_WO_Fill' in master_df.columns:
                    master_df = master_df.drop(columns=['Temp_WO_Fill'])
            else:
                master_df = pd.concat([df_track.reset_index(drop=True), df_plan.reset_index(drop=True)], axis=1)
        else:
            master_df = df_track

        master_df = master_df.fillna("")

        st.subheader("📊 معاينة التقرير الموحد النهائي:")
        st.dataframe(master_df, use_container_width=True, height=600)
        
        # 4. تصدير وتنسيق ملف الإكسيل الاحترافي
        output_filename = "Master_Final_Knitting_Report.xlsx"
        
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            master_df.to_excel(writer, index=False, sheet_name='Master Report')
            
        wb = openpyxl.load_workbook(output_filename)
        ws = wb.active
        
        # تنسيق الهيدر (عريض بخلفية زرقاء مميزة)
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
                label="📥 تحميل التقرير النهائي المعتمد (Excel)",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ تم تصدير التقرير النهائي بالهيكل النموذجي المطلوب بنجاح تام!")
            
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
else:
    st.warning("⚠️ يرجى رفع ملف التراك الأساسي وملف البلان للبدء.")