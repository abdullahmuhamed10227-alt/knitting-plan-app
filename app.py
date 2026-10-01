import io
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(
    page_title="Smart Knitting Plan & Tracking System", page_icon="🧵", layout="wide"
)

st.title("🧵 النظام الذكي المدمج لخطط التريكو وتتبع الماكينات (Plan & Tracking)")
st.write(
    "قم برفع **ملف خطة الإنتاج (PDF)** وملف **تتبع الماكينات (Excel)** معاً ليقوم"
    " النظام بمعالجة البلان من A إلى L، وسحب حالة الماكينات والأوردرات الفعلية"
    " والمتبقي والإيكو وطول الغرزة في الأعمدة من M إلى R بدقة متناهية."
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


# 📌 واجهة رفع الملفين معاً في نفس الشاشة
col_up1, col_up2 = st.columns(2)
with col_up1:
  plan_file = st.file_uploader("1️⃣ اختر ملف خطة الإنتاج الأساسي (PDF)", type=["pdf"])
with col_up2:
  track_file = st.file_uploader("2️⃣ اختر ملف تتبع الماكينات الفعلي (Excel)", type=["xlsx", "xls", "csv"])

if plan_file is not None and track_file is not None:
  st.success("✅ تم رفع الملفين بنجاح! جاري المعالجة والدمج والمطابقة الذكية...")

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

        # 📌 الخطوة الثانية: قراءة ملف التراك (شيت OVER VIEW) واستخراج البيانات للأعمدة M-R
        if track_file.name.endswith(".csv"):
          df_track = pd.read_csv(track_file)
        else:
          xls = pd.ExcelFile(track_file)
          sheet_to_use = "OVER VIEW" if "OVER VIEW" in xls.sheet_names else xls.sheet_names[0]
          for s in xls.sheet_names:
            if s.strip().upper() == "OVER VIEW":
              sheet_to_use = s
              break
          df_track = pd.read_excel(track_file, sheet_name=sheet_to_use, header=None)

        track_headers = [str(c).strip() for c in df_track.iloc[0].values]
        df_track_data = df_track.iloc[1:].copy()
        df_track_data.columns = track_headers

        # توريث أوردرات التشغيل رأسياً (ffill) لمنع أي خلايا فارغة
        if "Work ORDER" in df_track_data.columns:
          if isinstance(df_track_data["Work ORDER"], pd.DataFrame):
            df_track_data.iloc[:, 2] = df_track_data.iloc[:, 2].ffill()
          else:
            df_track_data["Work ORDER"] = df_track_data["Work ORDER"].ffill()

        # بناء قاموس التتبع للماكينات بدقة حسب هيكل الشيت
        track_lookup = {}
        for _, row in df_track_data.iterrows():
          m_id_raw = ""
          for val in row.values:
            val_str = str(val).strip()
            if val_str.startswith("M") and len(val_str) <= 6:
              m_id_raw = val_str
              break
          if not m_id_raw and len(row.values) > 4:
            m_id_raw = str(row.values[4]).strip()

          if m_id_raw and m_id_raw != "nan":
            m_clean = "".join(m_id_raw.split()).upper()
            row_vals = list(row.values)
            
            wo_val = str(row_vals[2]) if len(row_vals) > 2 and pd.notna(row_vals[2]) else ""
            stitch_val = str(row_vals[12]) if len(row_vals) > 12 and pd.notna(row_vals[12]) else ""
            eco_val = str(row_vals[13]) if len(row_vals) > 13 and pd.notna(row_vals[13]) else ""
            rem_val = str(row_vals[19]) if len(row_vals) > 19 and pd.notna(row_vals[19]) else ""
            on_off_val = str(row_vals[24]) if len(row_vals) > 24 and pd.notna(row_vals[24]) else "ON"

            track_lookup[m_clean] = {
                "OnOff": "ON" if on_off_val.strip().upper() == "ON" else "OFF",
                "Remaining": rem_val if rem_val != "nan" else "",
                "WorkOrder": wo_val if wo_val != "nan" else "",
                "StitchLength": stitch_val if stitch_val != "nan" else "",
                "SampleNo": eco_val if eco_val != "nan" else ""
            }

        # 📌 الخطوة الثالثة: إضافة الأعمدة الجديدة من M إلى R بناءً على رقم الماكينة
        onoff_col, rem_col, wo_mach_col, stitch_col, sample_mach_col = [], [], [], [], []

        for _, row in df_sorted.iterrows():
          raw_m = str(row.get("Machine", "")).strip()
          m_clean = "".join(raw_m.split()).upper()

          if m_clean in track_lookup:
            data = track_lookup[m_clean]
            onoff_col.append(data["OnOff"])
            rem_col.append(data["Remaining"])
            wo_mach_col.append(data["WorkOrder"])
            stitch_col.append(data["StitchLength"])
            sample_mach_col.append(data["SampleNo"])
          else:
            onoff_col.append("OFF")
            rem_col.append("-")
            wo_mach_col.append("-")
            stitch_col.append("-")
            sample_mach_col.append("-")

        df_sorted["OnOff Machin"] = onoff_col
        df_sorted["Remaining Machin"] = rem_col
        df_sorted["Work Order"] = wo_mach_col
        df_sorted["Stitch Length"] = stitch_col
        df_sorted["Sampleng"] = sample_mach_col

        st.success(
            f"✅ تم دمج البلان مع الأعمدة المخصصة للتراك (M-R) بنجاح تام! | إجمالي السطور: **{len(df_sorted)}** صف"
        )
        st.write("### 📊 معاينة الجدول النهائي المدمج:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # 📌 تصدير إلى Excel مع التنسيق ودمج خلايا A-L
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Master_Corrected_Plan"
          )
          workbook = writer.book
          worksheet = workbook.sheets["Master_Corrected_Plan"]

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
            if "description" in col_name or "yarn" in col_name:
              worksheet.column_dimensions[col_letter].width = 40
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

        # 📌 تصدير إلى PDF مرتب ومنضبط
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
            fontSize=6.5,
            leading=8,
            alignment=1,
        )
        style_header = ParagraphStyle(
            name="Header_Table",
            fontName="Helvetica-Bold",
            fontSize=7.5,
            leading=9,
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
        col_widths = [810 / num_cols] * num_cols

        pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
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
              label="📥 تحميل الملف النهائي المدمج (Excel)",
              data=excel_data,
              file_name="Master_Corrected_Plan_With_Tracking.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )
        with col2:
          st.download_button(
              label="📥 تحميل التقرير (PDF)",
              data=pdf_data,
              file_name="Master_Corrected_Plan_With_Tracking.pdf",
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
  st.info("الرجاء رفع **الملفين معاً** (ملف خطة الإنتاج PDF + ملف تتبع الماكينات Excel) للبدء.")