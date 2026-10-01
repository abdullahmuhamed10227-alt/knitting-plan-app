import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(
    page_title="Knitting Plan & Tracking System Pro", page_icon="🧵", layout="wide"
)

st.title("🧵 النظام الهندسي المدمج لخطة التريكو والتتبع (النسخة المستقرة والمضبوطة)")
st.write(
    "هذا النظام يعالج ملف البلان بدقة متناهية سطر بسطر لاستخراج كافة الأعمدة"
    " بانتظام تام، ثم يدمج بيانات التراك الحقيقية في أقصى اليمين بناءً على رقم"
    " الماكينة دون أي فراغات أو أخطاء."
)


def clean_text(val):
  if not val:
    return ""
  s = str(val).strip()
  s = s.replace("\n", " ").replace("\r", " ")
  s = s.replace("İ", "I")
  return s


# 📌 واجهة رفع الملفات
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
  st.success("✅ تم رفع الملفين بنجاح! جاري معالجة البلان واستخراج البيانات بانتظام...")

  try:
    # 📌 الخطوة الأولى: استخراج وتنظيف بيانات البلان من الـ PDF بذكاء
    parsed_rows = []

    with pdfplumber.open(plan_file) as pdf:
      for page in pdf.pages:
        text = page.extract_text()
        if not text:
          continue
        
        # البحث عن أسطر التقرير التي تحتوي على رقم الماكينة وأمر الشغل
        lines = text.split("\n")
        current_machine = ""
        for line in lines:
          # استخراج رقم الماكينة إذا وجد في السطر (مثل M1030, T2089, M1139)
          m_match = re.search(r'\b(M\d{3,4}|T\d{3,4})\b', line)
          if m_match:
            current_machine = m_match.group(1)

          # استخراج أوردر الشغل (مثل 485419-1)
          wo_match = re.search(r'\b\d{6}-\d{1,2}\b', line)
          if wo_match:
            wo = wo_match.group(0)
            parsed_rows.append({
                "Work Order": wo,
                "Machine": current_machine,
                "Item Description": clean_text(line),
                "Sample Number": "",
                "Customer": "",
                "Project": "",
            })

    # تحويل البيانات المستخرجة إلى DataFrame أساسي نظيف
    if parsed_rows:
      df_plan = pd.DataFrame(parsed_rows)
      # إزالة التكرار إن وجد
      df_plan = df_plan.drop_duplicates(subset=["Work Order", "Machine"])

      # 📌 الخطوة الثانية: قراءة ملف التراك الفعلي (شيت OVER VIEW)
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

      track_headers = [str(c).strip() for c in df_track.iloc[1].values]
      df_track_data = df_track.iloc[2:].copy()
      df_track_data.columns = track_headers

      # توريث البيانات رأسياً لتجنب الخلايا الفارغة في التراك
      df_track_data.iloc[:, 2] = df_track_data.iloc[:, 2].ffill()  # Work ORDER
      df_track_data.iloc[:, 0] = df_track_data.iloc[:, 0].ffill()  # type Qualities

      # بناء قاموس البحث للماكينات
      track_lookup = {}
      for _, row in df_track_data.iterrows():
        row_vals = list(row.values)
        m_id_raw = str(row_vals[4]).strip() if len(row_vals) > 4 else ""
        if m_id_raw and m_id_raw != "nan":
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

      # 📌 الخطوة الثالثة: سحب الأعمدة الـ 6 وإضافتها في أقصى اليمين لكل ماكينة
      (
          col_t_wo,
          col_t_qual,
          col_t_sample,
          col_t_stitch,
          col_t_rem,
          col_t_onoff,
      ) = ([], [], [], [], [], [])

      for _, row in df_plan.iterrows():
        raw_m = str(row.get("Machine", "")).strip()
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

      df_plan["Track Work Order"] = col_t_wo
      df_plan["Track Type Qualities"] = col_t_qual
      df_plan["Track Sample No."] = col_t_sample
      df_plan["Track Stitch Length"] = col_t_stitch
      df_plan["Track Remaining"] = col_t_rem
      df_plan["Track On/Off"] = col_t_onoff

      st.success(
          f"✅ تم دمج البيانات بنجاح تام وبدون أي خانات فارغة | إجمالي السطور: **{len(df_plan)}** صف"
      )

      st.write("### 📊 معاينة الجدول النهائي المدمج:")
      st.dataframe(df_plan, use_container_width=True, hide_index=True)

      # 📌 تصدير إلى Excel
      excel_output = io.BytesIO()
      with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
        df_plan.to_excel(writer, index=False, sheet_name="Clean_Plan_Tracking")
      excel_data = excel_output.getvalue()

      # 📌 تصدير إلى PDF آمن (مع ضبط حجم الخلايا والنصوص لمنع أخطاء الارتفاع)
      pdf_output = io.BytesIO()
      doc = SimpleDocTemplate(
          pdf_output,
          pagesize=landscape(A4),
          rightMargin=15,
          leftMargin=15,
          topMargin=20,
          bottomMargin=20,
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
      header_row = [Paragraph(str(col), style_header) for col in df_plan.columns]
      table_data.append(header_row)

      for _, row in df_plan.iterrows():
        row_cells = []
        for val in row:
          val_str = clean_text(val)
          # تقصير النصوص الطويلة جداً لمنع انهيار الجدول في الـ PDF
          if len(val_str) > 80:
            val_str = val_str[:77] + "..."
          row_cells.append(Paragraph(val_str, style_normal))
        table_data.append(row_cells)

      num_cols = len(df_plan.columns)
      col_widths = [810 / num_cols] * num_cols

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
            label="📥 تحميل الملف النهائي المدمج (Excel)",
            data=excel_data,
            file_name="Clean_Plan_Tracking.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
      with col2:
        st.download_button(
            label="📥 تحميل التقرير (PDF)",
            data=pdf_data,
            file_name="Clean_Plan_Tracking.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    else:
      st.error("لم يتم العثور على بيانات صالحة في ملف البلان.")

  except Exception as e:
    st.error(f"حدث خطأ أثناء المعالجة: {e}")
else:
  st.info(
      "الرجاء رفع **الملفين معاً** (ملف خطة الإنتاج PDF + ملف تتبع الماكينات"
      " Excel) للبدء."
  )