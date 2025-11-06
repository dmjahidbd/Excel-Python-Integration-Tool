"""
Excel-Python Integration Tool - Advanced Version
Complete data analysis, visualization, AI insights, and predictive analytics

Author: Jahid Hassan
GitHub: github.com/dmjahidbd/Excel-Python-Integration-Tool
"""
import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objects as go
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from scipy import stats
from excel_python_tool import ExcelPythonTool
from datetime import datetime

st.set_page_config(page_title="Excel-Python Tool - Advanced", page_icon="📊", layout="wide")

st.markdown("""<style>
.main-header {font-size: 3rem; font-weight: bold; color: #1f77b4; text-align: center; margin-bottom: 1rem;}
.sub-header {font-size: 1.2rem; text-align: center; color: #666; margin-bottom: 2rem;}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📊 Excel-Python Integration Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">🚀 Advanced Analytics | AI Insights | Predictive Models</div>', unsafe_allow_html=True)

with st.sidebar:
    st.image("https://img.shields.io/badge/Python-3.8%2B-blue", use_container_width=True)
    st.markdown("### 🎯 Features")
    st.markdown("- 📊 Multiple Charts\n- 🔥 Heatmaps\n- 🤖 AI Insights\n- 📈 Predictions\n- 💬 Chat Data\n- 📄 Export PDF/CSV")
    st.markdown("---\n### 👨‍💻 Developer\n**Jahid Hassan**\n[GitHub](https://github.com/dmjahidbd)")

if 'df' not in st.session_state:
    st.session_state.df = None
if 'insights' not in st.session_state:
    st.session_state.insights = []

def generate_ai_insights(df):
    insights = []
    try:
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            if len(df) > 1:
                values = df[col].dropna()
                if len(values) > 1:
                    slope = np.polyfit(range(len(values)), values, 1)[0]
                    if abs(slope) > values.std() * 0.1:
                        direction = "increasing" if slope > 0 else "decreasing"
                        insights.append(f"🔍 {col} is {direction}")
        for col in numeric_cols:
            Q1, Q3 = df[col].quantile(0.25), df[col].quantile(0.75)
            IQR = Q3 - Q1
            outliers = df[(df[col] < Q1 - 1.5 * IQR) | (df[col] > Q3 + 1.5 * IQR)]
            if len(outliers) > 0:
                insights.append(f"⚠️ {col} has {len(outliers)} outliers")
        if len(numeric_cols) > 1:
            corr_matrix = df[numeric_cols].corr()
            for i in range(len(corr_matrix.columns)):
                for j in range(i+1, len(corr_matrix.columns)):
                    if abs(corr_matrix.iloc[i, j]) > 0.7:
                        insights.append(f"🔗 {corr_matrix.columns[i]} & {corr_matrix.columns[j]} correlated")
    except:
        pass
    return insights if insights else ["✅ Data looks healthy!"]

def create_prediction(df, target, features):
    try:
        X = df[features].fillna(df[features].mean())
        y = df[target].fillna(df[target].mean())
        model = LinearRegression().fit(X, y)
        return model.predict(X), model.score(X, y)
    except:
        return None, 0

def chat_data(df, query):
    q = query.lower()
    try:
        nums = df.select_dtypes(include=[np.number]).columns
        if 'mean' in q or 'average' in q:
            return "📊 Averages:\n" + "\n".join([f"- {c}: {df[c].mean():.2f}" for c in nums])
        elif 'max' in q:
            return "📈 Maximum:\n" + "\n".join([f"- {c}: {df[c].max():.2f}" for c in nums])
        elif 'min' in q:
            return "📉 Minimum:\n" + "\n".join([f"- {c}: {df[c].min():.2f}" for c in nums])
        elif 'count' in q:
            return f"📊 Rows: {len(df)}, Columns: {len(df.columns)}"
        else:
            return "🤔 Ask about: mean, max, min, or count"
    except:
        return "⚠️ Error processing query"

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📤 Upload", "📊 Visualizations", "🤖 AI Insights", "📈 Predictions", "💬 Chat", "💾 Export"])

with tab1:
    st.header("Upload Excel File")
    uploaded = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    if uploaded:
        try:
            df = pd.read_excel(uploaded)
            st.session_state.df = df
            st.success(f"✅ Loaded {len(df)} rows, {len(df.columns)} columns")
            st.dataframe(df.head(20), use_container_width=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Rows", len(df))
            c2.metric("Columns", len(df.columns))
            c3.metric("Numeric", len(df.select_dtypes(include=[np.number]).columns))
            c4.metric("Missing", df.isnull().sum().sum())
        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("👆 Upload file or try sample data")
        if st.button("🔄 Sample Data"):
            st.session_state.df = pd.DataFrame({
                'Product': ['A', 'B', 'C'] * 10,
                'Region': ['East', 'West'] * 15,
                'Sales': np.random.randint(100, 300, 30),
                'Quantity': np.random.randint(10, 30, 30),
                'Profit': np.random.randint(20, 50, 30)
            })
            st.rerun()

with tab2:
    st.header("📊 Visualizations")
    if st.session_state.df is not None:
        df = st.session_state.df
        nums = df.select_dtypes(include=[np.number]).columns.tolist()
        cats = df.select_dtypes(include=['object']).columns.tolist()
        
        viz = st.selectbox("Chart Type", ["Bar", "Line", "Pie", "Scatter", "Histogram", "Box", "Heatmap", "Area"])
        
        if viz == "Bar" and nums:
            x = st.selectbox("X", df.columns, key="bx")
            y = st.selectbox("Y", nums, key="by")
            if st.button("Generate"):
                st.plotly_chart(px.bar(df, x=x, y=y, title=f"{y} by {x}"), use_container_width=True)
        
        elif viz == "Line" and nums:
            cols = st.multiselect("Columns", nums)
            if cols and st.button("Generate"):
                st.plotly_chart(px.line(df, y=cols, title="Trends"), use_container_width=True)
        
        elif viz == "Pie" and nums:
            cat = st.selectbox("Category", df.columns)
            val = st.selectbox("Values", nums)
            if st.button("Generate"):
                data = df.groupby(cat)[val].sum().reset_index()
                st.plotly_chart(px.pie(data, names=cat, values=val), use_container_width=True)
        
        elif viz == "Scatter" and len(nums) >= 2:
            x = st.selectbox("X", nums, key="sx")
            y = st.selectbox("Y", nums, key="sy")
            if st.button("Generate"):
                st.plotly_chart(px.scatter(df, x=x, y=y, title=f"{y} vs {x}"), use_container_width=True)
        
        elif viz == "Histogram" and nums:
            col = st.selectbox("Column", nums)
            bins = st.slider("Bins", 10, 100, 30)
            if st.button("Generate"):
                st.plotly_chart(px.histogram(df, x=col, nbins=bins), use_container_width=True)
        
        elif viz == "Box" and nums:
            col = st.selectbox("Column", nums)
            if st.button("Generate"):
                st.plotly_chart(px.box(df, y=col), use_container_width=True)
        
        elif viz == "Heatmap" and len(nums) > 1:
            if st.button("Generate Correlation Heatmap"):
                fig, ax = plt.subplots(figsize=(10, 8))
                sns.heatmap(df[nums].corr(), annot=True, cmap='coolwarm', center=0, ax=ax)
                st.pyplot(fig)
        
        elif viz == "Area" and nums:
            cols = st.multiselect("Columns", nums, key="area")
            if cols and st.button("Generate"):
                st.plotly_chart(px.area(df, y=cols), use_container_width=True)
    else:
        st.info("Upload data first")

with tab3:
    st.header("🤖 AI Insights")
    if st.session_state.df is not None:
        if st.button("🔍 Generate AI Insights"):
            with st.spinner("Analyzing..."):
                insights = generate_ai_insights(st.session_state.df)
                st.session_state.insights = insights
        
        if st.session_state.insights:
            st.subheader("📋 Discovered Insights:")
            for insight in st.session_state.insights:
                st.markdown(f"- {insight}")
    else:
        st.info("Upload data first")

with tab4:
    st.header("📈 Predictive Analytics")
    if st.session_state.df is not None:
        df = st.session_state.df
        nums = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if len(nums) >= 2:
            target = st.selectbox("Target Variable", nums)
            features = st.multiselect("Features", [c for c in nums if c != target])
            
            if features and st.button("🚀 Train Model"):
                preds, score = create_prediction(df, target, features)
                if preds is not None:
                    st.success(f"✅ Model R² Score: {score:.3f}")
                    df_pred = df.copy()
                    df_pred['Predicted'] = preds
                    st.plotly_chart(px.scatter(df_pred, x=target, y='Predicted', 
                                             title="Actual vs Predicted"), use_container_width=True)
                else:
                    st.error("❌ Prediction failed")
        else:
            st.info("Need at least 2 numeric columns")
    else:
        st.info("Upload data first")

with tab5:
    st.header("💬 Chat with Your Data")
    if st.session_state.df is not None:
        query = st.text_input("Ask about your data:", placeholder="What is the average?")
        if query:
            response = chat_data(st.session_state.df, query)
            st.markdown(response)
        
        st.markdown("**Try asking:**")
        st.markdown("- What is the mean?\n- Show max values\n- Count rows")
    else:
        st.info("Upload data first")

with tab6:
    st.header("💾 Export & Reports")
    if st.session_state.df is not None:
        df = st.session_state.df
        
        st.subheader("📥 Download Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button("📄 Download CSV", csv, "data.csv", "text/csv")
        
        with col2:
            excel_buffer = io.BytesIO()
            df.to_excel(excel_buffer, index=False)
            st.download_button("📊 Download Excel", excel_buffer.getvalue(), "data.xlsx")
        
        with col3:
            txt = df.to_string()
            st.download_button("📝 Download TXT", txt, "data.txt", "text/plain")
        
        st.subheader("📊 Data Summary")
        st.write(df.describe())
        
    else:
        st.info("Upload data first")

st.markdown("---")
st.markdown('<div style="text-align: center; color: #666;"><p>Made with ❤️ by Jahid Hassan | <a href="https://github.com/dmjahidbd/Excel-Python-Integration-Tool">GitHub</a></p></div>', unsafe_allow_html=True)
