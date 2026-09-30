import io
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(
    page_title="Knitting Plan & Tracking System", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة خطط وتتبع ماكينات التريكو (Knitting Plan & Tracking)")
st.write(
    "هذا النظام موحد لرفع خطة الإنتاج (PDF) أو ملف تتبع الماكينات (Excel)"
    " ومعالجتها واختيار الأعمدة والتنسيق النهائي بدقة."
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
    "اختر الملف (ملف البلان PDF أو ملف تتبع الماكينات Excel)",
    type=["pdf", "xlsx", "csv"],
)

if uploaded_file is not None:
  file_name = uploaded_file.name.lower()
  st.success(f"تم رفع الملف ({uploaded_file.name}) بنجاح! جاري المعالجة...")

  try:
    # 📌 الحالة الأولى: لو الملف المرفوع هو ملف تتبع الماكينات (Excel Tracking)
    if file_name.endswith((".xlsx", ".xls", ".csv")):
      if file_name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
      else:
        xls = pd.ExcelFile(uploaded_file)
        sheet_to_use = xls.sheet_names[0]
        for s in xls.sheet_names:
          if "tracking" in s.lower():
            sheet_to_use = s
            break
        df = pd.read_excel(uploaded_file, sheet_name=sheet_to_use)

        # البحث عن الهيدر الصحيح في ملف التراك
        for idx, row in df.head(10).iterrows():
          row_str = " ".join(row.astype(str).values).lower()
          if (
              "mac. no" in row_str
              or "machine no" in row_str
              or "running workorder" in row_str
          ):
            df.columns = row.values
            df = df.iloc[idx + 1 :].reset_index(drop=True)
            break

      df.columns = [
          str(c).strip()
          for c in df.columns
          if pd.notna(c) and str(c).strip() != "nan"
      ]
      df = df.dropna(how="all")

      for col in df.columns:
        df[col] = df[col].apply(clean_text_single_line)

      # استخراج الأعمدة الرئيسية في التراك
      mach_col = next(
          (
              c
              for c in df.columns
              if "mac" in c.lower() or "machine" in c.lower()
          ),
          df.columns[0] if len(df.columns) > 0 else None,
      )
      wo_col = next(
          (
              c
              for c in df.columns
              if "workorder" in c.lower()
              or "work order" in c.lower()
              or "running" in c.lower()
          ),
          None,
      )
      cust_col = next(
          (c for c in df.columns if "customer" in c.lower()), None
      )
      fabric_col = next(
          (c for c in df.columns if "fabric" in c.lower() or "code" in c.lower()),
          None,
      )
      yarn_col = next(
          (c for c in df.columns if "yarn" in c.lower() or "mix" in c.lower()),
          None,
      )

      if mach_col:
        df_sorted = df.sort_values(by=mach_col).reset_index(drop=True)

        # إضافة حالة الماكينة وتفاصيلها في أخر الأعمدة على اليمين كما طلبت تماماً
        def determine_status(row):
          val_wo = str(row.get(wo_col, "")).strip() if wo_col else ""
          row_text = " ".join([str(v) for v in row.values]).lower()
          if (
              "closed" in row_text
              or "stop" in row_text
              or val_wo == ""
              or val_wo == "nan"
          ):
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

        df_sorted["Detailed_Status_Summary"] = df_sorted.apply(
            create_summary_details, axis=1
        )

        st.success(
            "✅ تم معالجة ملف تتبع الماكينات وإضافة تقرير الحالة والمواصفات في"
            f" أخر الشيت | إجمالي الماكينات: **{len(df_sorted)}**"
        )
        st.write("### 📊 معاينة جدول التتبع المحدث:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # تصدير إكسيل للتراك المحدث
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
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

      else:
        st.error("لم يتم العثور على عمود رقم الماكينة في ملف التراك.")

    # 📌 الحالة الثانية: لو الملف المرفوع هو خطة الإنتاج (PDF Plan)
    elif file_name.endswith(".pdf"):
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

        df.columns = [str(col).strip() for col in df.columns]

        # حذف الأعمدة غير المطلوبة
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
        seq_col = next(
            (col for col in df.columns if "seq" in col.lower()), None
        )
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
          sort_cols = [wo_col]
          if seq_col:
            sort_cols.append(seq_col)
          if "Machine" in df.columns:
            sort_cols.append("Machine")

          df_sorted = df.sort_values(by=sort_cols)

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
          final_columns_order = priority_cols + other_cols
          df_sorted = df_sorted[final_columns_order]

          st.success(
              "✅ تم بنجاح معالجة خطة الإنتاج (PDF) وترتيب الهيكل وإضافة الدمج |"
              f" إجمالي السطور: **{len(df_sorted)}** صف"
          )
          st.write("### 📊 معاينة خطة الإنتاج المنسقة:")
          st.dataframe(df_sorted, use_container_width=True, hide_index=True)

          # تصدير إكسيل خطة الإنتاج مع الدمج
          excel_output = io.BytesIO()
          with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
            df_sorted.to_excel(
                writer, index=False, sheet_name="Master_Corrected_Plan"
            )
            workbook = writer.book
            worksheet = workbook.active

            from openpyxl.styles import Alignment, Font, PatternFill

            header_fill = PatternFill(
                start_color="1F4E78", end_color="1F4E78", fill_type="solid"
            )
            header_font = Font(
                name="Calibri", size=11, bold=True, color="FFFFFF"
            )

            for col_num, col in enumerate(
                worksheet.iter_cols(min_row=1, max_row=1), 1
            ):
              for cell in col:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

            for col in worksheet.columns:
              col_letter = col[0].column_letter
              col_name = str(col[0].value).lower()
              if "description" in col_name or "yarn" in col_name:
                worksheet.column_dimensions[col_letter].width = 45
              elif "note" in col_name:
                worksheet.column_dimensions[col_letter].width = 25
              else:
                worksheet.column_dimensions[col_letter].width = 16

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

          # تصدير PDF لخطة الإنتاج
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
              fontSize=7,
              leading=9,
              alignment=1,
          )
          style_header = ParagraphStyle(
              name="Header_Table",
              fontName="Helvetica-Bold",
              fontSize=8,
              leading=10,
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
            for col_name, val in zip(df_sorted.columns, row):
              val_str = str(val) if pd.notna(val) else ""
              val_str = val_str.replace("İ", "I")
              row_cells.append(Paragraph(val_str, style_normal))
            table_data.append(row_cells)

          num_cols = len(df_sorted.columns)
          col_widths = []
          for col_name in df_sorted.columns:
            c_low = str(col_name).lower()
            if "description" in c_low or "yarn" in c_low:
              col_widths.append(110)
            elif "note" in c_low:
              col_widths.append(70)
            else:
              col_widths.append(
                  (810 - 250) / (num_cols - 2) if num_cols > 2 else 60
              )

          pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
          pdf_table.setStyle(
              TableStyle([
                  ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                  ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                  ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                  ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                  ("TOPPADDING", (0, 0), (-1, -1), 4),
                  ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
              ])
          )

          elements.append(pdf_table)
          doc.build(elements)
          pdf_data = pdf_output.getvalue()

          col1, col2 = st.columns(2)
          with col1:
            st.download_button(
                label="📥 تحميل خطة الإنتاج (Excel منسق ومدمج)",
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
          st.error("لم يتم العثور على عمود الـ Work Order في ملف الـ PDF.")

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع الملف (سواء خطة الإنتاج PDF أو ملف التتبع Excel) للبدء.")