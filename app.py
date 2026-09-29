import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وإعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "هذا التطبيق مخصص لترتيب تقارير الإنتاج وتصحيح وتعديل اتجاهات النصوص"
    " والأرقام المعكوسة لجميع الماكينات والجوجات تلقائياً."
)


# دالة عامة وجذرية لعكس أي نص معكوس (سواء جوج أو ماكينة) وإعادته لأصله السليم
def fix_reversed_text(val):
  if pd.isna(val):
    return val
  s = str(val).strip()
  if s:
    return s[::-1]
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

      # تطبيق دالة التصحيح الشاملة على أعمدة الماكينات والجوجات
      if "Machine" in df.columns:
        df["Machine"] = df["Machine"].apply(fix_reversed_text)

      if "Gauge_Specs" in df.columns:
        df["Gauge_Specs"] = df["Gauge_Specs"].apply(fix_reversed_text)

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

        st.write("### 📊 معاينة البيانات بعد التصحيح الشامل لكل الجوجات والماكينات:")
        # تم إخفاء الـ Index الافتراضي هنا بناءً على طلبك
        st.dataframe(df_sorted, use_container_width=True, hide_index=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_sorted.to_excel(
              writer, index=False, sheet_name="Master_Corrected_Plan"
          )
        processed_data = output.getvalue()

        st.download_button(
            label="📥 تحميل الملف النهائي المعتمد بالكامل (Excel)",
            data=processed_data,
            file_name="Master_Corrected_Knitting_Plan.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
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