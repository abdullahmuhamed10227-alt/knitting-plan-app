import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Master Knitting & Tracking System", page_icon="🧵", layout="wide"
)

st.title("🧵 النظام الهندسي المتكامل لربط البلان بالتراك (القاعدة الشاملة للماكينات)")
st.write(
    "يعتمد هذا النظام على دمج ملف البلان مع القاعدة الشاملة للماكينات في التراك،"
    " لضمان ظهور كافة الماكينات (الشاغلة والواقفة) مع بياناتها بدقة تامة."
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
  st.success("✅ تم رفع الملفين بنجاح! جاري مطابقة الماكينات ودمج البيانات...")

  try:
    # 📌 1. استخراج بيانات البلان من الـ PDF
    plan_data = []
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
              plan_data.append(cleaned)

    # 📌 2. قراءة ملف التراك الشامل (OVER VIEW)
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
      df_track = pd.read_excel(track_file, sheet_name=sheet_to_use, header=None)

    track_headers = [str(c).strip() for c in df_track.iloc[1].values]
    df_track_data = df_track.iloc[2:].copy()
    df_track_data.columns = track_headers

    df_track_data.iloc[:, 2] = df_track_data.iloc[:, 2].ffill()  # Work ORDER
    df_track_data.iloc[:, 0] = df_track_data.iloc[:, 0].ffill()  # type Qualities

    # بناء قاموس التراك الشامل لكل ماكينة
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

        track_lookup[m_clean] = {
            "Machine": m_id_raw,
            "Track Work Order": t_wo if t_wo != "nan" else "",
            "Track Type Qualities": t_qual if t_qual != "nan" else "",
            "Track Sample No.": t_sample if t_sample != "nan" else "",
            "Track Stitch Length": t_stitch if t_stitch != "nan" else "",
            "Track Remaining": t_rem if t_rem != "nan" else "",
            "Track On/Off": "ON" if t_onoff.upper() == "ON" else "OFF",
        }

    # تحويل الماكينات الشاملة من التراك إلى جدول أساسي، ودمج بيانات البلان عليها
    master_rows = []
    import re

    # استخراج معلومات البلان في ديكشنري مفهرس برقم الماكينة أو أوردر الشغل
    plan_dict = {}
    for p_row in plan_data:
      p_text = " ".join(p_row)
      m_match = re.search(r"\b(M\d{3,4}|T\d{3,4})\b", p_text)
      if m_match:
        m_key = "".join(m_match.group(1).split()).upper()
        plan_dict[m_key] = p_row

    # دمج القائمتين بناءً على الماكينات الموجودة في التراك
    for m_clean, t_data in track_lookup.items():
      combined_row = t_data.copy()
      if m_clean in plan_dict:
        p_row = plan_dict[m_clean]
        combined_row["Plan Work Order"] = p_row[1] if len(p_row) > 1 else ""
        combined_row["Item Description"] = p_row[3] if len(p_row) > 3 else ""
        combined_row["Sample Number"] = p_row[4] if len(p_row) > 4 else ""
        combined_row["Customer Name"] = p_row[11] if len(p_row) > 11 else ""
      else:
        combined_row["Plan Work Order"] = ""
        combined_row["Item Description"] = "ماكينة متوقفة / بدون خطة حالية"
        combined_row["Sample Number"] = ""
        combined_row["Customer Name"] = ""
      master_rows.append(combined_row)

    df_master = pd.DataFrame(master_rows)

    st.success(
        f"✅ تم دمج قاعدة الماكينات الشاملة بنجاح تام! | إجمالي الماكينات في"
        f" النظام: **{len(df_master)}** ماكينة"
    )

    st.write("### 📊 معاينة الجدول الشامل المدمج:")
    st.dataframe(df_master, use_container_width=True, hide_index=True)

    # تصدير إلى Excel
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
      df_master.to_excel(
          writer, index=False, sheet_name="Master_Machines_Tracking"
      )
    excel_data = excel_output.getvalue()

    st.download_button(
        label="📥 تحميل الشامل المدمج (Excel)",
        data=excel_data,
        file_name="Master_Machines_Tracking.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

  except Exception as e:
    st.error(f"حدث خطأ أثناء المعالجة: {e}")
else:
  st.info(
      "الرجاء رفع الملفين معاً (البلان PDF + التراك Excel) لاعتماد القاعدة"
      " الشاملة."
  )