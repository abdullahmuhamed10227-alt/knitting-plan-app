import pandas as pd
import glob
import os
import streamlit as st

st.title("نظام تخطيط ومتابعة التريكو الديناميكي")

def generate_master_knitting_report():
    # البحث التلقائي عن ملف التراك في مجلد المشروع بدلاً من استخدام اسم ثابت
    excel_files = glob.glob('*.xlsx')
    tracking_files = [f for f in excel_files if 'TRACKING' in f or 'M..C' in f]
    
    if not tracking_files:
        raise FileNotFoundError("لم يتم العثور على ملف التراك (Tracking Excel File) في المجلد!")
    
    # اختيار الملف المتاح
    tracking_filepath = tracking_files[0]
    print(f"جاري استخدام ملف التراك: {tracking_filepath}")
    st.info(f"جاري استخدام ملف التراك: {tracking_filepath}")
    
    xls = pd.ExcelFile(tracking_filepath)
    
    # قراءة شيت التتبع الأساسي (OVER VIEW)
    df_overview = pd.read_excel(tracking_filepath, sheet_name='OVER VIEW', header=1)
    df_overview = df_overview.dropna(how='all')
    
    # قراءة شيت أوردرات التسلسل (F.K.G)
    try:
        df_fkg = pd.read_excel(tracking_filepath, sheet_name='F.K.G')
    except Exception:
        df_fkg = pd.DataFrame()
        
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
        
    return master_df

try:
    master_report = generate_master_knitting_report()
    st.success("تم معالجة وتوليد التقرير النهائي بنجاح!")
    st.dataframe(master_report.head(20))
    
    # زر تحميل الملف النهائي
    output_filename = "Master_Knitting_Report.xlsx"
    master_report.to_excel(output_filename, index=False)
    
    with open(output_filename, "rb") as file:
        st.download_button(
            label="تحميل التقرير النهائي (Excel)",
            data=file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except Exception as e:
    st.error(f"حدث خطأ أثناء التنفيذ: {e}")