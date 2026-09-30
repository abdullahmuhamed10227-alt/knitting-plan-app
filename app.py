import io
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وإعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "هذا التطبيق مخصص لترتيب تقارير الإنتاج، تنقية الأعمدة، دمج خلايا أمر"
    " الشغل، وإضافة تقرير حالة وتتبع الماكينات في نهاية الشيت."
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
    "اختر ملف خطة الإنتاج أو التتبع (PDF أو Excel)", type=["pdf", "xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة البيانات وتحديث الأعمدة...")

  try:
    df = None
    if uploaded_file.name.endswith(".xlsx"):
      xls = pd.ExcelFile(uploaded_file)
      # قراءة الشيت الأول افتراضياً
      df = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0])

      # التحقق إذا كان ملف تتبع (Tracking) يحتوي على صف عنوان مكرر أو هيدر في الصف الأول
      if "Unnamed" in str(df.columns[1]) or df.iloc[0].astype(str).str.contains("Work ORDER|Machine NO").any():
        # إعادة تعيين الهيدر من الصف الأول
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)

    elif uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)

    elif uploaded_file.name.endswith(".pdf"):
      all_rows = []
      headers = None

      with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
          table = page.extract_table()
          if table:
            if headers is None:
              headers = table[0]
              rows = table[1:]
            else:
              rows = [
                  row
                  for row in table[1:]
                  if row != headers
                  and not any("Work Order" in str(cell) for cell in row)
              ]
            all_rows.extend(rows)

      if all_rows and headers:
        clean_headers = []
        seen = {}
        for h in headers:
          h_str = str(h).strip() if h else ""
          if h_str in seen:
            seen[h_str] += 1
            clean_headers.append(f"{h_str}_{seen[h_str]}")
          else:
            seen[h_str] = 0
            clean_headers.append(h_str if h_str else f"Col_{seen[h_str]}")

        df = pd.DataFrame(all_rows, columns=clean_headers)

    if df is not None and not df.empty:
      df.columns = [str(col).strip() for col in df.columns]

      # حذف أعمدة Start, Finish, Acs إن وجدت لعدم الحاجة إليها
      cols_to_drop = []
      for col in df.columns:
        col_lower = col.lower()
        if (
            "start" in col_lower
            or "finish" in col_lower
            or col_lower in ["acs.", "acs"]
        ):
          cols_to_drop.append(col)
      if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

      # تنظيف النصوص
      for col in df.columns:
        if "machine" in col.lower() or "gauge" in col.lower():
          df[col] = df[col].apply(fix_reversed_text)
        else:
          df[col] = df[col].apply(clean_text_single_line)

      # البحث عن أعمدة الـ Work Order والـ Machine
      wo_col = next(
          (
              col
              for col in df.columns
              if "work order" in col.lower() or "work_order" in col.lower()
          ),
          None,
      )
      mach_col = next(
          (
              col
              for col in df.columns
              if "machine" in col.lower() or "machine no" in col.lower()
          ),
          None,
      )
      seq_col = next((col for col in df.columns if "seq" in col.lower()), None)
      sample_col = next(
          (col for col in df.columns if "sample" in col.lower()), None
      )
      cust_col = next(
          (col for col in df.columns if "customer" in col.lower() or "custmer" in col.lower()), None
      )
      proj_col = next(
          (col for col in df.columns if "project" in col.lower()), None
      )

      if wo_col:
        sort_cols = [wo_col]
        if seq_col:
          sort_cols.append(seq_col)
        if mach_col and mach_col in df.columns:
          sort_cols.append(mach_col)

        df_sorted = df.sort_values(by=sort_cols)

        # توحيد البيانات المشتركة لكل Work Order
        shared_cols_to_merge = []
        if sample_col:
          shared_cols_to_merge.append(sample_col)
        if cust_col:
          shared_cols_to_merge.append(cust_col)
        if proj_col:
          shared_cols_to_merge.append(proj_col)

        for col in shared_cols_to_merge:
          df_sorted[col] = df_sorted.groupby(wo_col)[col].transform(
              lambda x: x.iloc[0] if not x.empty else ""
          )

        # 📌 إضافة عمود أخير في نهاية الشيت يوضح حالة الماكينة وتفاصيلها (Machine Summary & Status)
        if mach_col:
          def get_machine_summary(row):
            m_name = row.get(mach_col, "Unknown")
            w_ord = row.get(wo_col, "N/A")
            status = "شغالة (Running)" if pd.notna(w_ord) and str(w_ord).strip() != "" else "واقفة (Stopped)"
            return f"مكينة: {m_name} | الحالة: {status} | أوردر: {w_ord}"

          df_sorted["Machine_Status_Summary"] = df_sorted.apply(get_machine_summary, axis=1)

        total_rows = len(df_sorted)
        st.success(
            f"✅ تم معالجة الملف بنجاح وإضافة تقرير حالة الماكينات في نهاية الشيت | إجمالي السطور:"
            f" **{total_rows}** صف"
        )

        st.write("### 📊 معاينة البيانات بعد التحديث:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # تصدير إلى Excel مع دمج الخلايا وتنسيق الأعمدة
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Master_Tracking_Plan"
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
            if "description" in col_name or "yarn" in col_name or "summary" in col_name:
              worksheet.column_dimensions[col_letter].width = 40
            elif "note" in col_name:
              worksheet.column_dimensions[col_letter].width = 25
            else:
              worksheet.column_dimensions[col_letter].width = 16

            for cell in col:
              cell.alignment = Alignment(
                  vertical="center", horizontal="center", wrap_text=True
              )

          # دمج خلايا Work Order والبيانات المشتركة رأسياً
          cols_to_merge_indices = []
          for idx, col_name in enumerate(df_sorted.columns, start=1):
            if col_name == wo_col or col_name in shared_cols_to_merge:
              cols_to_merge_indices.append(idx)

          if cols_to_merge_indices:
            start_row = 2
            while start_row <= worksheet.max_row:
              wo_val = worksheet.cell(row=start_row, column=1).value
              end_row = start_row
              while end_row + 1 <= worksheet.max_row:
                if worksheet.cell(row=end_row + 1, column=1).value == wo_val:
                  end_row += 1
                else:
                  break

              if end_row > start_row:
                for col_idx in cols_to_merge_indices:
                  worksheet.merge_cells(
                      start_row=start_row,
                      start_column=col_idx,
                      end_row=end_row,
                      end_column=col_idx,
                  )
              start_row = end_row + 1

        excel_data = excel_output.getvalue()

        # زر التنزيل
        st.download_button(
            label="📥 تحميل ملف التتبع المحدث والنهائي (Excel)",
            data=excel_data,
            file_name="Updated_Machine_Tracking_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

      else:
        st.error("لم يتم العثور على عمود الـ Work Order في البيانات.")
        st.dataframe(df, hide_index=True)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع ملف التتبع (Excel أو PDF) للبدء.")