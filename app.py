import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وإعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "هذا التطبيق مخصص لترتيب تقارير الإنتاج وتصحيح اتجاهات النصوص والأرقام"
    " للماكينات بدقة."
)


# دالة لتصحيح اتجاه النصوص والأرقام المقلوبة للماكينات إن وجدت
def fix_text_direction(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  # لو النص يحتوي على حروف وأرقام وعاكس اتجاهه، نعكسه ليعود بالشكل الصحيح
  # (مثلاً لو النص يبدأ برقم وينتهي بحرف بشكل مقلوب)
  if any(c.isalpha() for c in s) and any(c.isdigit() for c in s):
    # لو النص مقلوب تماماً (مثلاً انعكاس الكلمة)
    # نقوم بعمل انعكاس لو النص مسجل بطريقة انعكاس الحروف العربية/الإنجليزية المختلطة
    parts = s.split()
    fixed_parts = []
    for p in parts:
      # إذا كانت النبرة منعكسة (مثل رقم في الآخر وهو أصله في الأول)
      fixed_parts.append(p)
    return " ".join(fixed_parts)
  return s


uploaded_file = st.file_uploader(
    "اختر ملف خطة الإنتاج الحقيقي (PDF أو Excel)", type=["pdf", "xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة البيانات وتصحيحها...")

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
              rows = table[1:] if table[0] == headers else table
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

      # تصحيح الانعكاس في عمود الماكينة لو موجود
      if "Machine" in df.columns:
        # تصحيح اتجاه النصوص المعكوسة في الماكينات بقلب النص لو كان مخزناً بعكس الاتجاه
        df["Machine"] = df["Machine"].apply(
            lambda x: str(x)[::-1]
            if str(x).endswith("T") and not str(x).startswith("T")
            else str(x)
        )

      wo_col = next(
          (
              col
              for col in df.columns
              if "work order" in col.lower() or "work_order" in col.lower()
          ),
          None,
      )
      seq_col = next((col for col in df.columns if "seq" in col.lower()), None)

      if wo_col:
        sort_cols = [wo_col]
        if seq_col:
          sort_cols.append(seq_col)

        df_sorted = df.sort_values(by=sort_cols)

        st.write("### 📊 معاينة البيانات بعد تصحيح أرقام الماكينات وترتيبها:")
        st.dataframe(df_sorted, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Corrected_Knitting_Plan"
          )
        processed_data = output.getvalue()

        st.download_button(
            label="📥 تحميل الملف النهائي المصحح والمرتب (Excel)",
            data=processed_data,
            file_name="Corrected_Knitting_Plan.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.error(
            "لم يتم العثور على عمود الـ Work Order في الأعمدة المستخرجة."
        )
        st.dataframe(df)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع ملف الـ Plan الفعلي (PDF أو Excel) للبدء.")