import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="داشبورد البيانات", layout="wide")

st.title("📊 داشبورد البيانات")
st.markdown("ارفع ملف Excel أو CSV وشوف بياناتك تتحول لداشبورد تلقائياً")

# ── رفع الملف ──────────────────────────────────────────────
uploaded_file = st.file_uploader("ارفع الملف هنا", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # قراءة الملف
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"✅ تم تحميل الملف — {df.shape[0]} صف، {df.shape[1]} عمود")

    # ── معاينة البيانات ──────────────────────────────────────
    with st.expander("📋 معاينة البيانات", expanded=True):
        st.dataframe(df, use_container_width=True)

    # ── الأعمدة الرقمية والنصية ─────────────────────────────
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols     = df.select_dtypes(include="object").columns.tolist()

    # ── KPIs ────────────────────────────────────────────────
    if numeric_cols:
        st.subheader("📈 ملخص سريع")
        cols = st.columns(min(len(numeric_cols), 4))
        for i, col in enumerate(numeric_cols[:4]):
            cols[i].metric(col, f"{df[col].sum():,.0f}", f"متوسط: {df[col].mean():,.1f}")

    st.divider()

    # ── رسوم بيانية ─────────────────────────────────────────
    st.subheader("📊 الرسوم البيانية")

    col1, col2 = st.columns(2)

    with col1:
        if numeric_cols:
            y_col = st.selectbox("اختر العمود الرقمي", numeric_cols, key="y")
            if cat_cols:
                x_col = st.selectbox("اختر عمود التصنيف", cat_cols, key="x")
                fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} حسب {x_col}",
                             color=x_col, template="plotly_dark")
            else:
                fig = px.histogram(df, x=y_col, title=f"توزيع {y_col}", template="plotly_dark")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if len(numeric_cols) >= 2:
            x2 = st.selectbox("محور X", numeric_cols, key="x2")
            y2 = st.selectbox("محور Y", numeric_cols, index=1, key="y2")
            fig2 = px.scatter(df, x=x2, y=y2, title=f"{y2} مقابل {x2}",
                              template="plotly_dark")
            st.plotly_chart(fig2, use_container_width=True)
        elif cat_cols and numeric_cols:
            fig2 = px.pie(df, names=cat_cols[0], values=numeric_cols[0],
                          title=f"توزيع {numeric_cols[0]}", template="plotly_dark")
            st.plotly_chart(fig2, use_container_width=True)

else:
    st.info("⬆️ ارفع ملف Excel أو CSV للبداية")
