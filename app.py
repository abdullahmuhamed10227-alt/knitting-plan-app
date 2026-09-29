import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إدارة وإعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "هذا التطبيق مخصص لمهندسي الإنتاج لتحليل تقارير الـ Plan وعكس ترتيبها بحيث"
    " يكون **رقم أمر الشغل (Work Order)** هو الأساس أمام جميع الماكينات"
    " والكميات."
)

uploaded_file = st.file_uploader(
    "اختر ملف خطة الإنتاج الحقيقي (PDF أو Excel)", type=["pdf", "xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة واستخراج البيانات...")

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
          h_str = str(h).strip() if h else "col"
          if h_str in seen:
            seen[h_str] += 1
            clean_headers.append(f"{h_str}_{seen[h_str]}")
          else:
            seen[h_str] = 0
            clean_headers.append(h_str)

        df = pd.DataFrame(all_rows, columns=clean_headers)

        valid_cols = [
            c
            for c in df.columns
            if not c.startswith("col") and c != "None" and c != ""
        ]
        if valid_cols:
          df = df[valid_cols]

      else:
        st.warning(
            "لم يتم العثور على جداول واضحة داخل ملف الـ PDF. تأكد من تنسيق"
            " الملف."
        )

    if df is not None and not df.empty:
      df.columns = [str(col).strip() for col in df.columns]

      # البحث عن أعمدة أمر الشغل والماكينة
      wo_col = next(
          (
              col
              for col in df.columns
              if "Work Order" in col or "Work_Order" in col
          ),
          None,
      )
      mach_col = next(
          (col for col in df.columns if "Acs" in col or "Machine" in col), None
      )
      seq_col = next((col for col in df.columns if "Seq" in col), None)

      if wo_col:
        # تجميع وترتيب الجدول بحيث يتم جلب الـ Work Order أولاً،
        # وتحته مباشرة تتجمع كل الماكينات والـ Seq المرتبطة به.
        sort_cols = [wo_col]
        if seq_col:
          sort_cols.append(seq_col)
        if mach_col:
          sort_cols.append(mach_col)

        df_sorted = df.sort_values(by=sort_cols)

        st.write(
            "### 📊 معاينة البيانات بعد تجميع أمر الشغل مع كافة الماكينات:"
        )
        st.dataframe(df_sorted, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_sorted.to_excel(writer, index=False, sheet_name="WorkOrder_Centric")
        processed_data = output.getvalue()

        st.download_button(
            label="📥 تحميل الملف المرتب والمجمع حسب أمر الشغل (Excel)",
            data=processed_data,
            file_name="WorkOrder_Centric_Plan.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.error(
            "لم يتم العثور على عمود الـ Work Order بأسماء الأعمدة المستخرجة."
        )
        st.dataframe(df)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع ملف الـ Plan الفعلي (PDF أو Excel) للبدء.")