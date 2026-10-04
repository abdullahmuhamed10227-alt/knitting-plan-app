import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Master Knitting & Tracking System Pro", page_icon="🧵", layout="wide"
)

st.title("🧵 النظام الهندسي المتكامل لربط البلان بالتراك (النسخة المصححة والدقيقة)")
st.write(
    "هذا النظام يرتب الجدول بناءً على خطة الإنتاج (البلان) أولاً، ثم يدمج"
    " القاعدة الشاملة للماكينات من التراك بدقة تامة، مع احترام حالات التشغيل"
    " الفعلي (ON / OFF) لكل ماكينة بدون أي تداخل."
)


def clean_text(val):
  if pd.isna(val) or not val:
    return ""
  s = str(val).strip()
  s = s.replace("\n", " ").replace("\r", " ")
  s = s.replace("İ", "I")
  return s


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
  st.success("✅ تم رفع الملفين بنجاح! جاري معالجة وترتيب البيانات بدقة...")

  try:
    # 📌 1. استخراج بيانات البلان من الـ PDF
    plan_rows = []
    with pdfplumber.open(plan_file) as pdf:
      for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
          for row in table:
            cleaned = [clean_text(c) for c in row]
            if (
                any("Work Order" in c for c in cleaned)
                or any("Seq." in c for c in cleaned)
            ):
              continue
            if len(cleaned) >= 2 and cleaned[1]:
              plan_rows.append(cleaned)

    # 📌 2. قراءة ملف التراك الشامل (OVER VIEW) بدون ffill مدمر للحالات
    if track_file.name.endswith(".csv"):
      df_track = pd.read_csv(track_file)
    else:
      xls = pd.ExcelFile(track_file)
      sheet_to_use = (
          "OVER VIEW" if "OVER VIEW" in xls.sheet_names else xls.sheet_names[0]
      )
      for s in xls.sheet_names:
        if s.strip().upper() == "OVER VIEW":
          sheet_to_use = s
          break
      df_track = pd.read_excel(
          track_file, sheet_sheet=sheet_to_use, header=None
      )

    track_headers = [str(c).strip() for c in df_track.iloc[1].values]
    df_track_data = df_track.iloc[2:].copy()
    df_track_data.columns = track_headers

    # بناء قاموس التراك الفعلي لكل ماكينة بدقة
    track_lookup = {}
    for _, row in df_track_data.iterrows():
      row_vals = list(row.values)
      m_id_raw = str(row_vals[4]).strip() if len(row_vals) > 4 else ""
      if m_id_raw and m_id_raw != "nan":
        m_clean = "".join(m_id_raw.split()).upper()

        t_qual = (
            str(row_vals[0]).strip()
            if len(row_vals) > 0 and pd.notna(row_vals[0])
            else ""
        )
        t_wo = (
            str(row_vals[2]).strip()
            if len(row_vals) > 2 and pd.notna(row_vals[2])
            else ""
        )
        t_stitch = (
            str(row_vals[12]).strip()
            if len(row_vals) > 12 and pd.notna(row_vals[12])
            else ""
        )
        t_sample = (
            str(row_vals[13]).strip()
            if len(row_vals) > 13 and pd.notna(row_vals[13])
            else ""
        )
        t_rem = (
            str(row_vals[19]).strip()
            if len(row_vals) > 19 and pd.notna(row_vals[19])
            else ""
        )
        t_onoff = (
            str(row_vals[24]).strip()
            if len(row_vals) > 24 and pd.notna(row_vals[24])
            else "OFF"
        )

        # التصحيح الهام: إذا كانت الماكينة OFF أو ليس لها أوردر شغل في التراك، يترك الـ Work Order فارغاً تماماً
        is_on = t_onoff.upper() == "ON"
        final_wo = t_wo if (is_on and t_wo and t_wo != "nan") else ""

        track_lookup[m_clean] = {
            "Machine": m_id_raw,
            "Track Work Order": final_wo,
            "Track Type Qualities": t_qual if t_qual != "nan" else "",
            "Track Sample No.": t_sample if t_sample != "nan" else "",
            "Track Stitch Length": t_stitch if t_stitch != "nan" else "",
            "Track Remaining": t_rem if t_rem != "nan" else "",
            "Track On/Off": "ON" if is_on else "OFF",
        }

    # 📌 3. بناء الجدول المدمج بحيث نبدأ بترتيب البلان أولاً
    import re

    master_rows = []
    processed_machines = set()

    # معالجة ماكينات البلان أولاً لضمان ترتيبها بالأسلوب المتفق عليه
    for p_row in plan_rows:
      p_wo = p_row[1] if len(p_row) > 1 else ""
      p_desc = p_row[3] if len(p_row) > 3 else ""
      p_sample = p_row[4] if len(p_row) > 4 else ""
      p_cust = p_row[11] if len(p_row) > 11 else ""

      p_text = " ".join(p_row)
      m_match = re.search(r"\b(M\d{3,4}|T\d{3,4})\b", p_text)
      m_clean = (
          "".join(m_match.group(1).split()).upper() if m_match else "UNKNOWN"
      )
      m_raw = m_match.group(1) if m_match else ""

      # جلب بيانات التراك لهذه الماكينة إن وجدت
      t_data = track_lookup.get(
          m_clean,
          {
              "Machine": m_raw,
              "Track Work Order": "",
              "Track Type Qualities": "",
              "Track Sample No.": "",
              "Track Stitch Length": "",
              "Track Remaining": "",
              "Track On/Off": "ON",
          },
      )

      combined = {
          "Machine": m_raw if m_raw else m_clean,
          "Plan Work Order": p_wo,
          "Item Description": p_desc,
          "Sample Number": p_sample,
          "Customer Name": p_cust,
          "Track Work Order": t_data["Track Work Order"],
          "Track Type Qualities": t_data["Track Type Qualities"],
          "Track Sample No.": t_data["Track Sample No."],
          "Track Stitch Length": t_data["Track Stitch Length"],
          "Track Remaining": t_data["Track Remaining"],
          "Track On/Off": t_data["Track On/Off"],
      }
      master_rows.append(combined)
      if m_clean != "UNKNOWN":
        processed_machines.add(m_clean)

    # إضافة باقي ماكينات المصنع (التي لم تورد في البلان أو الواقفة OFF) لتكتمل القاعدة الشاملة
    for m_clean, t_data in track_lookup.items():
      if m_clean not in processed_machines:
        combined = {
            "Machine": t_data["Machine"],
            "Plan Work Order": "",
            "Item Description": (
                "ماكينة متوقفة / بدون خطة في البلان"
                if t_data["Track On/Off"] == "OFF"
                else "ماكينة عاملة / غير مدرجة في البلان الحالي"
            ),
            "Sample Number": "",
            "Customer Name": "",
            "Track Work Order": t_data["Track Work Order"],
            "Track Type Qualities": t_data["Track Type Qualities"],
            "Track Sample No.": t_data["Track Sample No."],
            "Track Stitch Length": t_data["Track Stitch Length"],
            "Track Remaining": t_data["Track Remaining"],
            "Track On/Off": t_data["Track On/Off"],
        }
        master_rows.append(combined)

    df_master = pd.DataFrame(master_rows)

    st.success(
        f"✅ تم ضبط الترتيب والحالات بنجاح تام! | إجمالي السطور المعروضة:"
        f" **{len(df_master)}** صف"
    )

    st.write("### 📊 معاينة الجدول النهائي المرتب والمصحح:")
    st.dataframe(df_master, use_container_width=True, hide_index=True)

    # تصدير إلى Excel
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
      df_master.to_excel(
          writer, index=False, sheet_name="Master_Ordered_Plan"
      )
    excel_data = excel_output.getvalue()

    st.download_button(
        label="📥 تحميل الملف النهائي المرتب (Excel)",
        data=excel_data,
        file_name="Master_Ordered_Plan.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

  except Exception as e:
    st.error(f"حدث خطأ أثناء المعالجة: {e}")
else:
  st.info("الرجاء رفع ملف البلان (PDF) وملف التراك (Excel) معاً للبدء.")