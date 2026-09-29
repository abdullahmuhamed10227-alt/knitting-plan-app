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

# رفع الملف (يدعم PDF و Excel)
uploaded_file = st.file_uploader(
    "اختر ملف خطة الإنتاج الحقيقي (PDF أو Excel)", type=["pdf", "xlsx", "csv"]
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة واستخراج البيانات...")

  try:
    df = None
    # 1. لو الملف Excel أو CSV
    if uploaded_file.name.endswith(".xlsx"):
      df = pd.read_excel(uploaded_file)
    elif uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)

    # 2. لو الملف PDF (استخراج الجداول الحقيقية من كل الصفحات)
    elif uploaded_file.name.endswith(".pdf"):
      all_data = []
      with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
          table = page.extract_table()
          if table:
            all_data.extend(table)

      if all_data:
        # أول صف يعتبر هو عناوين الأعمدة (Headers)
        headers = all_data[0]
        rows = all_data[1:]
        df = pd.DataFrame(rows, columns=headers)
      else:
        st.warning(
            "لم يتم العثور على جداول واضحة داخل ملف الـ PDF. يرجى التأكد من أن"
            " الملف يحتوي على جداول نصية."
        )

    if df is not None and not df.empty:
      # تنظيف أسماء الأعمدة لإزالة المسافات الزائدة
      df.columns = [str(col).strip() for col in df.columns]

      # البحث عن عمود أمر الشغل والترتيب بغض النظر عن المسافات
      wo_col = next(
          (col for col in df.columns if "Work Order" in col or "Work_Order" in col),
          None,
      )
      seq_col = next((col for col in df.columns if "Seq" in col), None)

      if wo_col:
        # فرز البيانات وترتيبها بناءً على أمر الشغل
        if seq_col:
          df_sorted = df.sort_values(by=[wo_col, seq_col])
        else:
          df_sorted = df.sort_values(by=[wo_col])

        st.write("### 📊 معاينة البيانات الحقيقية بعد إعادة الهيكلة:")
        st.dataframe(df_sorted, use_container_width=True)

        # زر تحميل الملف الناتج بصيغة Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_sorted.to_excel(writer, index=False, sheet_name="Restructured_Plan")
        processed_data = output.getvalue()

        st.download_button(
            label="📥 تحميل الملف الجديد بالكامل مرتباً (Excel)",
            data=processed_data,
            file_name="Real_Restructured_Knitting_Plan.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.error(
            "لم يتم العثور على عمود الـ Work Order في الملف المرفوع. تأكد من"
            " تطابق أسماء الأعمدة."
        )
        st.dataframe(df)

  except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
else:
  st.info("الرجاء رفع ملف الـ Plan الفعلي (PDF أو Excel) للبدء.")