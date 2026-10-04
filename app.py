import pandas as pd
import numpy as np

def generate_master_knitting_report(tracking_filepath, output_filepath='Master_Knitting_Report.xlsx'):
    """
    نموذج ذكي وموحد لدمج ملف التراك الأساسي مع أوردرات التسلسل (البلان) بشكل ديناميكي
    مع الحفاظ الكامل على كافة تفاصيل التراك دون حذف أو تغيير.
    """
    print("1. جاري قراءة وتحميل ملف التراك الأساسي...")
    xls = pd.ExcelFile(tracking_filepath)
    
    # قراءة شيت التتبع الأساسي (OVER VIEW) مع اعتماد الصف الثاني كروؤس للأعمدة
    df_overview = pd.read_excel(tracking_filepath, sheet_name='OVER VIEW', header=1)
    df_overview = df_overview.dropna(how='all') # إزالة الصفوف الفارغة تماماً
    
    print(f"   - عدد الماكينات في ملف التراك: {len(df_overview)}")

    # قراءة شيت أوردرات التسلسل (F.K.G)
    try:
        df_fkg = pd.read_excel(tracking_filepath, sheet_name='F.K.G')
        print("   - تم قراءة شيت الأوردرات والمتسلسلات (F.K.G) بنجاح.")
    except Exception as e:
        df_fkg = pd.DataFrame()
        print("   - تنبيه: لم يتم العثور على شيت F.K.G، سيتم الاكتفاء ببيانات التراك الأساسية.")

    # معالجة وتوزيع أوردرات البلان في أعمدة جديدة على اليمين (ح حسب السيكونس Sira)
    if not df_fkg.empty and 'Machine No' in df_fkg.columns and 'Sira' in df_fkg.columns:
        # تحويل جدول الأوردرات المتسلسلة إلى جدول محوري (Pivot Table) لتوزيع السيكونس بجانب بعضه
        df_fkg_pivot = df_fkg.pivot_table(
            index='Machine No',
            columns='Sira',
            values=['Work Order Name', 'is Emri Kalemi', 'Fabric Code', 'Planned Qty', 'Remained Qty', 'Release Date'],
            aggfunc='first'
        )
        
        # توحيد وتسمية الأعمدة الجديدة الناتجة عن السيكونس بوضوح (مثال: Seq_10_Work Order Name)
        df_fkg_pivot.columns = [f"Seq_{s}_{val}" for val, s in df_fkg_pivot.columns]
        df_fkg_pivot = df_fkg_pivot.reset_index()
        
        # البحث عن اسم عمود الماكينة في ملف التراك لربطه بدقة
        mc_col_candidates = [c for c in df_overview.columns if 'machine' in str(c).lower() or 'Machine NO' in str(c)]
        if mc_col_candidates:
            mc_col = mc_col_candidates[0]
            # دمج الجدول الأساسي للتراك مع أعمدة السيكونس الجديدة على اليمين
            master_df = pd.merge(df_overview, df_fkg_pivot, left_on=mc_col, right_on='Machine No', how='left')
        else:
            master_df = df_overview
    else:
        master_df = df_overview

    # تصدير الجدول النهائي المحدث إلى ملف Excel منظم
    master_df.to_excel(output_filepath, index=False)
    print(f"\n✅ تمت عملية الدمج وتوليد التقرير النهائي بنجاح!")
    print(f"   - إجمالي أعمدة الجدول النهائي: {len(master_df.columns)}")
    print(f"   - تم حفظ الملف باسم: {output_filepath}")
    
    return master_df

# تشغيل الدالة واستخراج التقرير الفوري
master_report = generate_master_knitting_report('M..C TRACKING 2026 (6)(1).xlsx')