# ترتيب البيانات هرمياً
        sort_cols = [wo_col]
        if seq_col:
          sort_cols.append(seq_col)
        if "Machine" in df.columns:
          sort_cols.append("Machine")

        df_sorted = df.sort_values(by=sort_cols)

        # 📌 توحيد البيانات المشتركة (Sample, Customer, Project) لكل Work Order لتظهر ككتلة واحدة موحدة
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