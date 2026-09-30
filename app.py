import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Machine Tracking & Status Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وتتبع حالة الماكينات (Machine Tracking Restructurer)")
st.write(
    "هذا التطبيق مخصص لرفع ملف تتبع الماكينات (MC Tracking)، تنظيمه، تصحيح"
    " اتجاهات النصوص، وإضافة تقرير حالة تشغيل الماكينة والأوردر والمواصفات في"
    " نهاية الشيت بدقة."
)


def fix_reversed_text(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  if s:
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("İ", "I")
    return s[::-1]
  return s


def clean_text_single_line(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  return s.replace("\n", " ").replace("\r", " ")


uploaded_file = st.file_uploader(
    "اختر ملف تتبع الماكينات الفعلي (Excel)", type=["xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع ملف التراك بنجاح! جاري معالجة البيانات وتحديث الأعمدة...")

  try:
    df = None
    if uploaded_file.name.endswith(".xlsx"):
      xls = pd.ExcelFile(uploaded_file)
      # البحث عن شيت التراك المناسب
      sheet_to_use = xls.sheet_names[0]
      for s in xls.sheet_names:
        if "tracking" in s.lower():
          sheet_to_use = s
          break

      df = pd.read_excel(uploaded_file, sheet_name=sheet_to_use)

      # تنظيف وتحديد الهيدر الصحيح إذا كان ملف تراك يحتوي على صفوف علوية
      for idx, row in df.head(10).iterrows():
        row_str = " ".join(row.astype(str).values).lower()
        if "mac. no" in row_str or "machine no" in row_str or "running workorder" in row_str:
          df.columns = row.values
          df = df.iloc[idx + 1 :].reset_index(drop=True)
          break

    elif uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)

    if df is not None and not df.empty:
      # تنظيف أسماء الأعمدة وحذف الأعمدة الفارغة تماماً
      df.columns = [str(c).strip() for c in df.columns if pd.notna(c) and str(c).strip() != "nan"]
      df = df.dropna(how="all")

      # تنظيف النصوص وعكسها إذا لزم الأمر
      for col in df.columns:
        df[col] = df[col].apply(clean_text_single_line)

      # البحث عن الأعمدة الأساسية (الماكينة، الأوردر، العميل، الخيوط، إلخ)
      mach_col = next(
          (c for c in df.columns if "mac" in c.lower() or "machine" in c.lower()),
          df.columns[0] if len(df.columns) > 0 else None
      )
      wo_col = next(
          (c for c in df.columns if "workorder" in c.lower() or "work order" in c.lower() or "running" in c.lower()),
          None
      )
      cust_col = next((c for c in df.columns if "customer" in c.lower()), None)
      fabric_col = next((c for c in df.columns if "fabric" in c.lower() or "code" in c.lower()), None)
      yarn_col = next((c for c in df.columns if "yarn" in c.lower() or "mix" in c.lower()), None)

      if mach_col:
        # ترتيب البيانات حسب رقم الماكينة
        df_sorted = df.sort_values(by=mach_col).reset_index(drop=True)

        # 📌 إضافة أعمدة الحالة والمواصفات في أخر الشيت على اليمين كما طلبت
        def determine_status(row):
          # فحص ما إذا كانت الماكينة شغالة أو واقفة بناءً على أوردر الشغل أو عمود الـ Fabric/Status
          val_wo = str(row.get(wo_col, "")).strip() if wo_col else ""
          row_text = " ".join([str(v) for v in row.values]).lower()
          
          if "closed" in row_text or "stop" in row_text or val_wo == "" or val_wo == "nan":
            return "واقفة (Stopped)"
          else:
            return "شغالة (Running)"

        df_sorted["Machine_Status"] = df_sorted.apply(determine_status, axis=1)

        def create_summary_details(row):
          m = row.get(mach_col, "N/A")
          w = row.get(wo_col, "لا يوجد أوردر") if wo_col else "N/A"
          c = row.get(cust_col, "-") if cust_col else "-"
          f = row.get(fabric_col, "-") if fabric_col else "-"
          y = row.get(yarn_col, "-") if yarn_col else "-"
          status = row["Machine_Status"]
          return f"الماكينة: {m} | الحالة: {status} | الأوردر: {w} | العميل: {c} | القماش: {f} | الغزل: {y}"

        df_sorted["Detailed_Status_Summary"] = df_sorted.apply(create_summary_details, axis=1)

        total_rows = len(df_sorted)
        st.success(
            f"✅ تم بنجاح معالجة ملف تتبع الماكينات وإضافة تقرير الحالة والمواصفات في أخر الشيت | إجمالي عدد الماكينات:"
            f" **{total_rows}** ماكينة"
        )

        st.write("### 📊 معاينة جدول التتبع المحدث:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # تجهيز ملف الإكسيل للتنزيل مع تنسيق الأعمده و Wrap Text
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Updated_MC_Tracking"
          )
          workbook = writer.book
          worksheet = workbook.active

          from openpyxl.styles import Alignment, Font, PatternFill

          header_fill = PatternFill(
              start_color="1F4E78", end_color="1F4E78", fill_type="solid"
          )
          header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

          for col_num, col in enumerate(worksheet.iter_cols(min_row=1, max_row=1), 1):
            for cell in col:
              cell.fill = header_fill
              cell.font = header_font
              cell.alignment = Alignment(
                  horizontal="center", vertical="center", wrap_text=True
              )

          for col in worksheet.columns:
            col_letter = col[0].column_letter
            col_name = str(col[0].value).lower()
            if "summary" in col_name or "yarn" in col_name or "mix" in col_name:
              worksheet.column_dimensions[col_letter].width = 45
            elif "status" in col_name:
              worksheet.column_dimensions[col_letter].width = 20
            else:
              worksheet.column_dimensions[col_letter].width = 16

            for cell in col:
              cell.alignment = Alignment(
                  vertical="center", horizontal="center", wrap_text=True
              )

        excel_data = excel_output.getvalue()

        st.download_button(
            label="📥 تحميل ملف تتبع الماكينات المحدث والنهائي (Excel)",
            data=excel_data,
            file_name="Updated_MC_Tracking_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

      else:
        st.error("لم يتم العثور على عمود رقم الماكينة في الملف.")
        st.dataframe(df, hide_index=True)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة ملف التتبع: {e}")
else:
  st.info("الرجاء رفع ملف تتبع الماكينات (MC Tracking Excel) للبدء.")