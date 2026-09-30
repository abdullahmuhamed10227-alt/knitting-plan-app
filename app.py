import io
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(
    page_title="Integrated Knitting Plan & Tracking System", page_icon="🧵", layout="wide"
)

st.title("🧵 النظام المتكامل لإدارة وتتبع خطط التريكو (Plan & Tracking Integration)")
st.write(
    "هذا النظام يقوم برفع **ملف البلان (PDF)** ومعالجته وترتيبه هرمياً، ثم مطابقة"
    " وجلب كافة تفاصيل **ملف التراك (شيت OVER VIEW بعد توريث الأوردرات)** لكل ماكينة بدقة تامة وبدون أي نقص."
)


def fix_reversed_text(val):
  if pd.isna(val):
    return ""
  s = str(val).strip()
  if s:
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("İ", "I")
    return s[::-1]
  return s


def clean_text_single_line(val):
  if pd.isna(val):
    return ""
  s = str(val).strip()
  return s.replace("\n", " ").replace("\r", " ")


# رفع الملفين معاً
col_up1, col_up2 = st.columns(2)
with col_up1:
  plan_file = st.file_uploader("1️⃣ اختر ملف خطة الإنتاج الأساسي (PDF)", type=["pdf"])
with col_up2:
  track_file = st.file_uploader("2️⃣ اختر ملف تتبع الماكينات الفعلي (Excel)", type=["xlsx", "xls", "csv"])

if plan_file is not None and track_file is not None:
  st.success("✅ تم رفع الملفين بنجاح! جاري الدمج والمطابقة الشاملة لكافة الأعمدة...")

  try:
    # 📌 الخطوة الأولى: معالجة واستخراج جدول البلان (PDF)
    all_rows = []
    headers = None

    with pdfplumber.open(plan_file) as pdf:
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

      df_plan = pd.DataFrame(all_rows, columns=clean_headers)

      cols_list = list(df_plan.columns)
      rename_dict = {}
      if len(cols_list) > 0:
        rename_dict[cols_list[0]] = "Gauge_Specs"
      if len(cols_list) > 1:
        rename_dict[cols_list[1]] = "Machine"

      if rename_dict:
        df_plan = df_plan.rename(columns=rename_dict)

      if "Machine" in df_plan.columns:
        df_plan["Machine"] = df_plan["Machine"].ffill()
      if "Gauge_Specs" in df_plan.columns:
        df_plan["Gauge_Specs"] = df_plan["Gauge_Specs"].ffill()

      df_plan.columns = [str(col).strip() for col in df_plan.columns]

      # حذف الأعمدة غير المطلوبة
      cols_to_drop = []
      for col in df_plan.columns:
        col_lower = col.lower()
        if (
            "start" in col_lower
            or "finish" in col_lower
            or col_lower in ["acs.", "acs"]
        ):
          cols_to_drop.append(col)
      if cols_to_drop:
        df_plan = df_plan.drop(columns=cols_to_drop)

      for col in df_plan.columns:
        if col in ["Machine", "Gauge_Specs"]:
          df_plan[col] = df_plan[col].apply(fix_reversed_text)
        else:
          df_plan[col] = df_plan[col].apply(clean_text_single_line)

      wo_col = next(
          (
              col
              for col in df_plan.columns
              if "work order" in col.lower() or "work_order" in col.lower()
          ),
          None,
      )
      seq_col = next(
          (col for col in df_plan.columns if "seq" in col.lower()), None
      )
      sample_col = next(
          (col for col in df_plan.columns if "sample" in col.lower()), None
      )
      cust_col = next(
          (col for col in df_plan.columns if "customer" in col.lower()), None
      )
      proj_col = next(
          (col for col in df_plan.columns if "project" in col.lower()), None
      )

      if wo_col:
        sort_cols = [wo_col, seq_col] if seq_col else [wo_col]
        if "Machine" in df_plan.columns:
          sort_cols.append("Machine")

        df_sorted = df_plan.sort_values(by=sort_cols)

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

        priority_cols = [wo_col]
        if "Machine" in df_sorted.columns:
          priority_cols.append("Machine")
        if seq_col and seq_col in df_sorted.columns:
          priority_cols.append(seq_col)
        if (
            "Gauge_Specs" in df_sorted.columns
            and "Gauge_Specs" not in priority_cols
        ):
          priority_cols.append("Gauge_Specs")

        other_cols = [
            c for c in df_sorted.columns if c not in priority_cols
        ]
        df_sorted = df_sorted[priority_cols + other_cols]

        # 📌 الخطوة الثانية: قراءة ملف التراك (شيت OVER VIEW) مع توريث الأوردرات (ffill)
        if track_file.name.endswith(".csv"):
          df_track = pd.read_csv(track_file)
        else:
          xls = pd.ExcelFile(track_file)
          sheet_to_use = "OVER VIEW" if "OVER VIEW" in xls.sheet_names else xls.sheet_names[0]
          for s in xls.sheet_names:
            if s.strip().upper() == "OVER VIEW":
              sheet_to_use = s
              break
          df_track = pd.read_excel(track_file, sheet_name=sheet_to_use)

        # ضبط هيدر شيت التراك من الصف الأول
        if df_track.iloc[0].astype(str).str.contains("Machine NO|Work ORDER").any():
          df_track.columns = [str(c).strip() for c in df_track.iloc[0].values]
          df_track = df_track.iloc[1:].reset_index(drop=True)
        else:
          df_track.columns = [str(c).strip() for c in df_track.columns]

        # توريث القيم المدمجة رأسياً في الإكسيل (مثل Work ORDER) لضمان عدم وجود خلايا فارغة
        df_track = df_track.ffill()

        track_mach_col = next((c for c in df_track.columns if "machine no" in c.lower() or c.lower() == "machine no."), df_track.columns[4] if len(df_track.columns) > 4 else None)

        track_data_dict = {}
        for _, row in df_track.iterrows():
          raw_m_id = str(row.get(track_mach_col, "")).strip()
          if raw_m_id and raw_m_id != "nan":
            m_id_clean = "".join(raw_m_id.split()).upper()
            row_dict = {str(k).strip(): (v if pd.notna(v) else "") for k, v in row.items() if pd.notna(k)}
            track_data_dict[m_id_clean] = row_dict

        track_columns = []
        if track_data_dict:
          first_key = list(track_data_dict.keys())[0]
          track_columns = [col for col in track_data_dict[first_key].keys() if col != track_mach_col]

        for t_col in track_columns:
          col_values = []
          for _, row in df_sorted.iterrows():
            raw_m_num = str(row.get("Machine", "")).strip()
            m_num_clean = "".join(raw_m_num.split()).upper()
            
            if m_num_clean in track_data_dict:
              val = track_data_dict[m_num_clean].get(t_col, "")
              col_values.append(val if val != "" else "-")
            else:
              col_values.append("غير متوفر بالتراك")
          df_sorted[f"Track_{t_col}"] = col_values

        st.success(
            f"✅ تم دمج خطة الإنتاج مع كافة تفاصيل وأعمدة ملف التراك (بعد توريث الأوردرات) بنجاح تـام! | إجمالي السطور: **{len(df_sorted)}** صف"
        )
        st.write("### 📊 معاينة الجدول النهائي المدمج بكافة التفاصيل:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # 📌 تصدير إلى Excel مع الدمج والتنسيق
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Integrated_Plan_Tracking"
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
            if "description" in col_name or "yarn" in col_name or "track" in col_name:
              worksheet.column_dimensions[col_letter].width = 25
            else:
              worksheet.column_dimensions[col_letter].width = 15

            for cell in col:
              cell.alignment = Alignment(
                  vertical="center", horizontal="center", wrap_text=True
              )

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

        # 📌 تصدير إلى PDF (الأعمدة الأساسية لمنع الضغط)
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
            fontSize=6,
            leading=7,
            alignment=1,
        )
        style_header = ParagraphStyle(
            name="Header_Table",
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=9,
            textColor=colors.whitesmoke,
            alignment=1,
        )

        pdf_cols = [c for c in df_sorted.columns if not c.startswith("Track_") or c in ["Track_Work ORDER", "Track_ON / OFF", "Track_Yarn 1 LOT"]]
        df_pdf = df_sorted[pdf_cols]

        table_data = []
        header_row = [
            Paragraph(str(col), style_header) for col in df_pdf.columns
        ]
        table_data.append(header_row)

        for _, row in df_pdf.iterrows():
          row_cells = []
          for col_name, val in zip(df_pdf.columns, row):
            val_str = str(val) if pd.notna(val) else ""
            val_str = val_str.replace("İ", "I")
            row_cells.append(Paragraph(val_str, style_normal))
          table_data.append(row_cells)

        num_cols = len(df_pdf.columns)
        col_widths = [810 / num_cols] * num_cols

        pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        pdf_table.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ])
        )

        elements.append(pdf_table)
        doc.build(elements)
        pdf_data = pdf_output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
          st.download_button(
              label="📥 تحميل التقرير الشامل النهائي (Excel - كامل التفاصيل)",
              data=excel_data,
              file_name="Integrated_Full_Knitting_Plan_Tracking.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )
        with col2:
          st.download_button(
              label="📥 تحميل التقرير المرتب (PDF)",
              data=pdf_data,
              file_name="Integrated_Full_Knitting_Plan_Tracking.pdf",
              mime="application/pdf",
              use_container_width=True,
          )

      else:
        st.error("لم يتم العثور على عمود الـ Work Order في ملف الـ PDF.")
    else:
      st.error("فشل في استخراج البيانات من ملف البلان.")

  except Exception as e:
    st.error(f"حدث خطأ أثناء المعالجة والدمج: {e}")
else:
  st.info("الرجاء رفع **الملفين معاً** (ملف خطة الإنتاج PDF + ملف تتبع الماكينات Excel) للبدء في الدمج والمعالجة الشاملة.")