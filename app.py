import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Knitting Plan Restructurer", page_icon="🧵", layout="wide"
)

st.title("🧵 نظام إعادة هيكلة خطة التريكو (Knitting Plan)")
st.write(
    "مرحباً بك يا بشمهندس! ارفع ملف الـ Plan الخاص بالإنتاج لتحويل وترتيب"
    " البيانات بحيث يكون **رقم أمر الشغل (Work Order)** هو الأساس."
)

uploaded_file = st.file_uploader(
    "اختر ملف خطة الإنتاج (Excel أو CSV أو محاكاة PDF)",
    type=["xlsx", "csv", "pdf"],
)

if uploaded_file is not None:
  st.success("تم رفع الملف بنجاح! جاري معالجة البيانات...")

  try:
    if uploaded_file.name.endswith(".xlsx"):
      df = pd.read_excel(uploaded_file)
    elif uploaded_file.name.endswith(".csv"):
      df = pd.read_csv(uploaded_file)
    else:
      # داتا تجريبية افتراضية في حالة ملف الـ PDF للتأكد من عمل الموقع فوراً
      data = {
          "Work_Order": ["485419-1", "486899-1", "486899-1", "489739-1"],
          "Machine": ["M1139", "M1041", "M1035", "M1041"],
          "Seq": [10, 10, 10, 20],
          "Pl_Tot_Qty": ["5136/6050", "770/2890", "803/2890", "1350/2700"],
          "Customer_Name": [
              "NIKE",
              "TOMMY HILFIGER",
              "TOMMY HILFIGER",
              "TOMMY HILFIGER",
          ],
      }
      df = pd.DataFrame(data)

    # الترتيب حسب أمر الشغل والـ Seq
    if "Work_Order" in df.columns and "Seq" in df.columns:
      df_sorted = df.sort_values(by=["Work_Order", "Seq"])
    else:
      df_sorted = df

    st.write("### 📊 معاينة البيانات بعد إعادة الهيكلة:")
    st.dataframe(df_sorted, use_container_width=True)

    # زر التحميل لملف الإكسيل
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_sorted.to_excel(writer, index=False, sheet_name="Restructured_Plan")
    processed_data = output.getvalue()

    st.download_button(
        label="📥 تحميل الملف الجديد مرتباً (Excel)",
        data=processed_data,
        file_name="Rearranged_Knitting_Plan.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"حدث خطأ أثناء المعالجة: {e}")
else:
    st.info("الرجاء رفع الملف للبدء.")