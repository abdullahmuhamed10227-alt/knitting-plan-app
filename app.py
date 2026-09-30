import io
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وإعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "هذا التطبيق مخصص لترتيب تقارير الإنتاج هرمياً وتصحيح وتعديل اتجاهات"
    " النصوص والأرقام المعكوسة وتنسيق ملفات الإكسيل والـ PDF بدقة."
)


def fix_reversed_text(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  if s:
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("İ", "I").replace("İ", "I")
    return s[::-1]
  return s


def clean_text_single_line(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  return s.replace("\n", " ").replace("\r", " ")


uploaded_file = st.file_uploader(
    "اختر ملف خطة الإنتاج الحقيقي (PDF أو Excel)", type=["pdf", "xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة البيانات وتعديل الهيكل...")

  try:
    df = None
    if uploaded_file.name.endswith(".xlsx"):
      df = pd.read_excel(uploaded_file)
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

        cols_list = list(df.columns)
        rename_dict = {}
        if len(cols_list) > 0:
          rename_dict[cols_list[0]] = "Gauge_Specs"
        if len(cols_list) > 1:
          rename_dict[cols_list[1]] = "Machine"

        if rename_dict:
          df = df.rename(columns=rename_dict)

        if "Machine" in df.columns:
          df["Machine"] = df["Machine"].ffill()
        if "Gauge_Specs" in df.columns:
          df["Gauge_Specs"] = df["Gauge_Specs"].ffill()

    if df is not None and not df.empty:
      df.columns = [str(col).strip() for col in df.columns]

      for col in df.columns:
        if col in ["Machine", "Gauge_Specs"]:
          df[col] = df[col].apply(fix_reversed_text)
        else:
          df[col] = df[col].apply(clean_text_single_line)

      wo_col = next(
          (
              col
              for col in df.columns
              if "work order" in col.lower() or "work_order" in col.lower()
          ),
          None,
      )
      seq_col = next((col for col in df.columns if "seq" in col.lower()), None)
      sample_col = next(
          (col for col in df.columns if "sample" in col.lower()), None
      )
      cust_col = next(
          (col for col in df.columns if "customer" in col.lower()), None
      )
      proj_col = next(
          (col for col in df.columns if "project" in col.lower()), None
      )

      if wo_col:
        # ترتيب هرمي دقيق حسب أمر الشغل ثم السيكونس ثم الماكينة
        sort_cols = [wo_col]
        if seq_col:
          sort_cols.append(seq_col)
        if "Machine" in df.columns:
          sort_cols.append("Machine")

        df_sorted = df.sort_values(by=sort_cols)

        # توحيد البيانات المشتركة (Sample, Customer, Project) لكل Work Order لتعكس الشكل المطلوب بالصورة
        if sample_col:
          df_sorted[sample_col] = df_sorted.groupby(wo_col)[
              sample_col
          ].transform(lambda x: x.iloc[0] if not x.empty else "")
        if cust_col:
          df_sorted[cust_col] = df_sorted.groupby(wo_col)[cust_col].transform(
              lambda x: x.iloc[0] if not x.empty else ""
          )
        if proj_col:
          df_sorted[proj_col] = df_sorted.groupby(wo_col)[proj_col].transform(
              lambda x: x.iloc[0] if not x.empty else ""
          )

        priority_cols = ["Gauge_Specs", "Machine"]
        if seq_col and seq_col in df_sorted.columns:
          priority_cols.append(seq_col)
        if wo_col and wo_col not in priority_cols:
          priority_cols.append(wo_col)

        other_cols = [
            c for c in df_sorted.columns if c not in priority_cols
        ]
        final_columns_order = priority_cols + other_cols
        df_sorted = df_sorted[final_columns_order]

        total_rows = len(df_sorted)
        st.success(
            f"✅ تم بنجاح استخراج وترتيب الهيكل الهرمي | إجمالي عدد السطور:"
            f" **{total_rows}** صف"
        )

        st.write("### 📊 معاينة البيانات بعد الهيكلة والضبط:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # تصدير إلى Excel
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Master_Corrected_Plan"
          )
          workbook = writer.book
          worksheet = writer.sheets["Master_Corrected_Plan"]

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
                  horizontal="center", vertical="center", wrap_text=False
              )

          for col in worksheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
              cell.alignment = Alignment(
                  vertical="center", horizontal="center", wrap_text=False
              )
              if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

        excel_data = excel_output.getvalue()

        # تصدير إلى PDF
        pdf_output = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_output,
            pagesize=landscape(A4),
            rightMargin=10,
            leftMargin=10,
            topMargin=15,
            bottomMargin=15,
        )
        elements = []

        style_normal = ParagraphStyle(
            name="Normal_Table",
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            alignment=1,
        )
        style_header = ParagraphStyle(
            name="Header_Table",
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=11,
            textColor=colors.whitesmoke,
            alignment=1,
        )

        table_data = []
        header_row = [
            Paragraph(str(col), style_header) for col in df_sorted.columns
        ]
        table_data.append(header_row)

        for _, row in df_sorted.iterrows():
          row_cells = []
          for val in row:
            val_str = str(val) if pd.notna(val) else ""
            val_str = val_str.replace("İ", "I")
            row_cells.append(Paragraph(val_str, style_normal))
          table_data.append(row_cells)

        num_cols = len(df_sorted.columns)
        col_width = 810 / num_cols if num_cols > 0 else 50

        pdf_table = Table(table_data, colWidths=[col_width] * num_cols, repeatRows=1)
        pdf_table.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ])
        )

        elements.append(pdf_table)
        doc.build(elements)
        pdf_data = pdf_output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
          st.download_button(
              label="📥 تحميل الملف النهائي (Excel)",
              data=excel_data,
              file_name="Master_Corrected_Knitting_Plan.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )
        with col2:
          st.download_button(
              label="📥 تحميل التقرير المرتب (PDF)",
              data=pdf_data,
              file_name="Master_Corrected_Knitting_Plan.pdf",
              mime="application/pdf",
              use_container_width=True,
          )

      else:
        st.error(
            "لم يتم العثور على عمود الـ Work Order في الأعمدة المستخرجة."
        )
        st.dataframe(df, hide_index=True)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع ملف الـ Plan الفعلي (PDF أو Excel) للبدء.")