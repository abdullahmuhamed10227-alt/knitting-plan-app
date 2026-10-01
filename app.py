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

st.title("🧵 النظام المدمج لخطة التريكو مع بيانات التراك الأصلية الخام")
st.write(
    "هذا النظام يحافظ على ملف البلان الأصلي سليماً تماماً بنسبة 100%، ويقوم بسحب"
    " أعمدة التراك الـ 6 (Work Order, Type Qualities, Sample, Stitch Length,"
    " Remaining, On/Off) بدقة متناهية من ملف التراك الحقيقي وإضافتها في أقصى"
    " اليمين بناءً على رقم الماكينة."
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


# 📌 واجهة رفع الملفين معاً
col_up1, col_up2 = st.columns(2)
with col_up1:
  plan_file = st.file_uploader(
      "1️⃣ اختر ملف خطة الإنتاج الأساسي (PDF)", type=["pdf"]
  )
with col_up2:
  track_file = st.file_uploader(
      "2️⃣ اختر ملف تتبع الماكينات الفعلي (Excel)", type=["xlsx", "xls", "csv"]
  )

if plan_file is not None and track_file is not None:
  st.success("✅ تم رفع الملفين بنجاح! جاري معالجة البلان وسحب بيانات التراك الحقيقية...")

  try:
    # 📌 الخطوة الأولى: معالجة البلان الأصلي تماماً كما اتفقنا وبدون أي مساس
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

        # 📌 الخطوة الثانية: قراءة ملف التراك الفعلي (شيت OVER VIEW) من ملف الإكسيل المرفوع
        if track_file.name.endswith(".csv"):
          df_track = pd.read_csv(track_file)
        else:
          xls = pd.ExcelFile(track_file)
          sheet_to_use = (
              "OVER VIEW"
              if "OVER VIEW" in xls.sheet_names
              else xls.sheet_names[0]
          )
          for s in xls.sheet_names:
            if s.strip().upper() == "OVER VIEW":
              sheet_to_use = s
              break
          df_track = pd.read_excel(
              track_file, sheet_name=sheet_to_use, header=None
          )

        # استخراج الهيدر من الصف رقم 1 والبيانات من الصف رقم 2 فصاعداً
        track_headers = [str(c).strip() for c in df_track.iloc[1].values]
        df_track_data = df_track.iloc[2:].copy()
        df_track_data.columns = track_headers

        # توريث أوردرات التشغيل وأنواع الجودة رأسياً (ffill) لأنها مدمجة رأسياً في شيت التراك
        df_track_data.iloc[:, 2] = df_track_data.iloc[:, 2].ffill()  # Work ORDER
        df_track_data.iloc[:, 0] = df_track_data.iloc[:, 0].ffill()  # type Qualities

        # بناء قاموس البحث للماكينات بدقة تامة من الملف الحقيقي
        track_lookup = {}
        for _, row in df_track_data.iterrows():
          row_vals = list(row.values)
          
          # عمود رقم الماكينة هو العمود رقم 4 (Machine NO.)
          m_id_raw = str(row_vals[4]).strip() if len(row_vals) > 4 else ""
          if m_id_raw and m_id_raw != "nan" and m_id_raw != "nan":
            m_clean = "".join(m_id_raw.split()).upper()

            t_qual = str(row_vals[0]).strip() if len(row_vals) > 0 and pd.notna(row_vals[0]) else ""
            t_wo = str(row_vals[2]).strip() if len(row_vals) > 2 and pd.notna(row_vals[2]) else ""
            t_stitch = str(row_vals[12]).strip() if len(row_vals) > 12 and pd.notna(row_vals[12]) else ""
            t_sample = str(row_vals[13]).strip() if len(row_vals) > 13 and pd.notna(row_vals[13]) else ""
            t_rem = str(row_vals[19]).strip() if len(row_vals) > 19 and pd.notna(row_vals[19]) else ""
            t_onoff = str(row_vals[24]).strip() if len(row_vals) > 24 and pd.notna(row_vals[24]) else "ON"

            track_lookup[m_clean] = {
                "Track_Work_Order": t_wo if t_wo != "nan" else "",
                "Track_Type_Qualities": t_qual if t_qual != "nan" else "",
                "Track_Sample_Eco": t_sample if t_sample != "nan" else "",
                "Track_Stitch_Length": t_stitch if t_stitch != "nan" else "",
                "Track_Remaining": t_rem if t_rem != "nan" else "",
                "Track_OnOff": "ON" if t_onoff.upper() == "ON" else "OFF",
            }

        # 📌 الخطوة الثالثة: سحب البيانات الخام الـ 6 لكل ماكينة وإضافتها في أقصى اليمين
        (
            col_t_wo,
            col_t_qual,
            col_t_sample,
            col_t_stitch,
            col_t_rem,
            col_t_onoff,
        ) = ([], [], [], [], [], [])

        for _, row in df_sorted.iterrows():
          raw_m = ""
          for col_name in df_sorted.columns:
            if "machine" in col_name.lower() or "mc" in col_name.lower():
              raw_m = str(row.get(col_name, "")).strip()
              break
          if not raw_m and len(row) > 1:
            raw_m = str(row.iloc[1]).strip()

          m_clean = "".join(raw_m.split()).upper()

          if m_clean in track_lookup:
            d = track_lookup[m_clean]
            col_t_wo.append(d["Track_Work_Order"])
            col_t_qual.append(d["Track_Type_Qualities"])
            col_t_sample.append(d["Track_Sample_Eco"])
            col_t_stitch.append(d["Track_Stitch_Length"])
            col_t_rem.append(d["Track_Remaining"])
            col_t_onoff.append(d["Track_OnOff"])
          else:
            col_t_wo.append("")
            col_t_qual.append("")
            col_t_sample.append("")
            col_t_stitch.append("")
            col_t_rem.append("")
            col_t_onoff.append("OFF")

        # إضافة الأعمدة الـ 6 في أقصى اليمين بنصوصها الأصلية الخام
        df_sorted["Track Work Order"] = col_t_wo
        df_sorted["Track Type Qualities"] = col_t_qual
        df_sorted["Track Sample No."] = col_t_sample
        df_sorted["Track Stitch Length"] = col_t_stitch
        df_sorted["Track Remaining"] = col_t_rem
        df_sorted["Track On/Off"] = col_t_onoff

        total_rows = len(df_sorted)
        st.success(
            f"✅ تم دمج البلان وسحب بيانات التراك الحقيقية (بقيمه الأصلية الكاملة بدون نقصان) بنجاح تام! | إجمالي السطور: **{total_rows}** صف"
        )

        st.write("### 📊 معاينة الجدول النهائي المدمج:")
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        # 📌 تصدير إلى Excel مع تنسيق ودمج خلايا Work Order للبلان الأصلي
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
            if "description" in col_name or "yarn" in col_name or "qualities" in col_name:
              worksheet.column_dimensions[col_letter].width = 35
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

        # 📌 تصدير إلى PDF مرتب
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
              file_name="Plan_With_Exact_Tracking_Data.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )
        with col2:
          st.download_button(
              label="📥 تحميل التقرير (PDF)",
              data=pdf_data,
              file_name="Plan_With_Exact_Tracking_Data.pdf",
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
  st.info(
      "الرجاء رفع **الملفين معاً** (ملف خطة الإنتاج PDF + ملف تتبع الماكينات"
      " Excel) للبدء."
  )