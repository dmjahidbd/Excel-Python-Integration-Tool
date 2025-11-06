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
from sklearn.preprocessing import StandardScaler
from scipy import stats
from excel_python_tool import ExcelPythonTool
import base64
from datetime import datetime
import json

# Page configuration
st.set_page_config(
    page_title="Excel-Python Integration Tool - Advanced",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        text-align: center;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .insight-box {
        background-color: #e3f2fd;
        padding: 15px;
        border-left: 5px solid #2196F3;
        border-radius: 5px;
        margin: 10px 0;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        background-color: #f0f2f6;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 Excel-Python Integration Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">🚀 Advanced Analytics | AI Insights | Predictive Models | Interactive Dashboards</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://img.shields.io/badge/Python-3.8%2B-blue", use_container_width=True)
    
    st.markdown("### 🎯 Advanced Features")
    st.markdown("""
    - 📊 **Multiple Chart Types**
    - 🔥 **Correlation Heatmaps**
    - 🤖 **AI Insights**
    - 📈 **Predictive Analytics**
    - 💬 **Chat with Data**
    - 📄 **Export PDF/CSV**
    - 🎨 **Interactive Dashboards**
    - 📊 **Pivot Analysis**
    - 📉 **Outlier Detection**
    - 🎯 **Trend Forecasting**
    """)
    
    st.markdown("---")
    st.markdown("### 📖 Quick Start")
    st.markdown("""
    1. Upload Excel file
    2. Explore visualizations
    3. Get AI insights
    4. Download reports
    """)
    
    st.markdown("---")
    st.markdown("### 👨‍💻 Developer")
    st.markdown("**Jahid Hassan**")
    st.markdown("[GitHub](https://github.com/dmjahidbd)")

# Initialize session state
if 'tool' not in st.session_state:
    st.session_state.tool = None
if 'df' not in st.session_state:
    st.session_state.df = None
if 'insights' not in st.session_state:
    st.session_state.insights = []
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

# Helper Functions
def generate_ai_insights(df):
    """Generate AI-powered insights from data"""
    insights = []
    
    try:
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        # Trend analysis
        for col in numeric_cols:
            if len(df) > 1:
                values = df[col].dropna()
                if len(values) > 1:
                    slope = np.polyfit(range(len(values)), values, 1)[0]
                    if abs(slope) > values.std() * 0.1:
                        direction = "increasing" if slope > 0 else "decreasing"
                        change_pct = (slope * len(values) / values.mean()) * 100
                        insights.append(f"🔍 **{col}** is {direction} by approximately {abs(change_pct):.1f}% over the dataset")
        
        # Outlier detection
        for col in numeric_cols:
            Q1 = df[col].quantile(0.25)
            Q3 = df[col].quantile(0.75)
            IQR = Q3 - Q1
            outliers = df[(df[col] < Q1 - 1.5 * IQR) | (df[col] > Q3 + 1.5 * IQR)]
            if len(outliers) > 0:
                pct = (len(outliers) / len(df)) * 100
                insights.append(f"⚠️ **{col}** has {len(outliers)} outliers ({pct:.1f}% of data) - values significantly outside normal range")
        
        # Statistical insights
        for col in numeric_cols:
            mean_val = df[col].mean()
            median_val = df[col].median()
            skew = df[col].skew()
            
            if abs(skew) > 1:
                skew_type = "right-skewed (high values)" if skew > 0 else "left-skewed (low values)"
                insights.append(f"📊 **{col}** is {skew_type} with skewness {skew:.2f}")
        
        # Correlation insights
        if len(numeric_cols) > 1:
            corr_matrix = df[numeric_cols].corr()
            high_corr = []
            for i in range(len(corr_matrix.columns)):
                for j in range(i+1, len(corr_matrix.columns)):
                    corr_val = corr_matrix.iloc[i, j]
                    if abs(corr_val) > 0.7:
                        high_corr.append((corr_matrix.columns[i], corr_matrix.columns[j], corr_val))
            
            for col1, col2, corr_val in high_corr:
                relationship = "positively correlated" if corr_val > 0 else "negatively correlated"
                strength = "very strongly" if abs(corr_val) > 0.9 else "strongly"
                insights.append(f"🔗 **{col1}** and **{col2}** are {strength} {relationship} (r={corr_val:.3f})")
        
        # Missing data insights
        missing = df.isnull().sum()
        total_missing = missing.sum()
        if total_missing > 0:
            for col in missing[missing > 0].index:
                pct = (missing[col] / len(df)) * 100
                if pct > 10:
                    insights.append(f"❗ **{col}** has significant missing data: {missing[col]} values ({pct:.1f}%)")
        
        # Data quality summary
        if total_missing == 0 and len(outliers) == 0:
            insights.append("✅ **Data Quality**: Excellent! No missing values or significant outliers detected")
        
    except Exception as e:
        insights.append(f"⚠️ Error generating insights: {str(e)}")
    
    return insights if insights else ["✅ Data looks healthy! No significant patterns detected."]

def create_prediction_model(df, target_col, feature_cols):
    """Create linear regression prediction model"""
    try:
        X = df[feature_cols].fillna(df[feature_cols].mean())
        y = df[target_col].fillna(df[target_col].mean())
        
        model = LinearRegression()
        model.fit(X, y)
        
        predictions = model.predict(X)
        score = model.score(X, y)
        
        # Feature importance
        importance = pd.DataFrame({
            'Feature': feature_cols,
            'Coefficient': model.coef_
        }).sort_values('Coefficient', ascending=False)
        
        return predictions, score, importance, model
    except Exception as e:
        return None, 0, None, None

def export_to_pdf(df, insights, charts_data=None):
    """Export analysis to PDF"""
    try:
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.lib.units import inch
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch)
        elements = []
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1f77b4'),
            spaceAfter=30,
            alignment=1
        )
        
        # Title
        elements.append(Paragraph("Excel Data Analysis Report", title_style))
        elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Data Summary
        elements.append(Paragraph("Data Summary", styles['Heading2']))
        summary_data = [
            ['Metric', 'Value'],
            ['Total Rows', f"{len(df):,}"],
            ['Total Columns', f"{len(df.columns)}"],
            ['Numeric Columns', f"{len(df.select_dtypes(include=[np.number]).columns)}"],
            ['Missing Values', f"{df.isnull().sum().sum():,}"],
            ['Memory Usage', f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB"]
        ]
        summary_table = Table(summary_data)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 20))
        
        # AI Insights
        elements.append(Paragraph("AI-Powered Insights", styles['Heading2']))
        for i, insight in enumerate(insights[:15], 1):
            clean_insight = insight.replace('**', '').replace('*', '')
            elements.append(Paragraph(f"{i}. {clean_insight}", styles['Normal']))
            elements.append(Spacer(1, 6))
        
        elements.append(Spacer(1, 20))
        
        # Statistical Summary
        elements.append(Paragraph("Statistical Summary", styles['Heading2']))
        desc = df.describe()
        if len(desc.columns) > 0:
            desc_data = [['Statistic'] + list(desc.columns[:5])]  # Limit to 5 columns
            for idx in desc.index:
                row = [idx] + [f"{desc.loc[idx, col]:.2f}" for col in desc.columns[:5]]
                desc_data.append(row)
            
            desc_table = Table(desc_data)
            desc_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(desc_table)
        
        doc.build(elements)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"PDF generation error: {str(e)}")
        return None

def chat_with_data(df, query):
    """Simple rule-based chat with data"""
    query = query.lower()
    response = ""
    
    try:
        # Basic statistics queries
        if 'mean' in query or 'average' in query:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                means = df[numeric_cols].mean()
                response = "📊 **Average Values:**\n\n"
                for col, val in means.items():
                    response += f"- **{col}**: {val:.2f}\n"
        
        elif 'max' in query or 'maximum' in query or 'highest' in query:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                maxs = df[numeric_cols].max()
                response = "📈 **Maximum Values:**\n\n"
                for col, val in maxs.items():
                    response += f"- **{col}**: {val:.2f}\n"
        
        elif 'min' in query or 'minimum' in query or 'lowest' in query:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                mins = df[numeric_cols].min()
                response = "📉 **Minimum Values:**\n\n"
                for col, val in mins.items():
                    response += f"- **{col}**: {val:.2f
