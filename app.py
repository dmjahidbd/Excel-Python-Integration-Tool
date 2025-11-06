"""Excel-Python Integration Tool - Streamlit Web App
Web interface for Excel data analysis and transformation
"""

import streamlit as st
import pandas as pd
import io
from excel_python_tool import ExcelPythonTool

# Page configuration
st.set_page_config(
    page_title="Excel-Python Integration Tool",
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
    .stDownloadButton {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 Excel-Python Integration Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your Excel files and perform advanced analysis with ease</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image("https://img.shields.io/badge/Python-3.8%2B-blue", use_container_width=True)
    st.markdown("### 🚀 Features")
    st.markdown("""
    - 📂 Upload Excel files
    - 📊 Statistical analysis
    - 🧹 Data cleaning
    - 🔢 Custom formulas
    - 📈 Aggregation & pivot
    - 💾 Download results
    """)
    
    st.markdown("---")
    st.markdown("### 📖 Instructions")
    st.markdown("""
    1. Upload your Excel file
    2. View basic statistics
    3. Clean and transform data
    4. Download processed file
    """)
    
    st.markdown("---")
    st.markdown("### 👨‍💻 Developer")
    st.markdown("**Jahid Hassan**")
    st.markdown("[GitHub](https://github.com/dmjahidbd)")

# Main content
tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload & Preview", "📊 Analysis", "🧹 Clean Data", "🔢 Transform"])

# Initialize session state
if 'tool' not in st.session_state:
    st.session_state.tool = None
if 'df' not in st.session_state:
    st.session_state.df = None

# Tab 1: Upload & Preview
with tab1:
    st.header("Upload Your Excel File")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload .xlsx or .xls files"
    )
    
    if uploaded_file is not None:
        try:
            # Read the file
            df = pd.read_excel(uploaded_file)
            st.session_state.df = df
            
            # Create tool instance
            st.session_state.tool = ExcelPythonTool()
            st.session_state.tool.df = df
            
            st.success(f"✅ File uploaded successfully! {len(df)} rows and {len(df.columns)} columns loaded.")
            
            # Display preview
            st.subheader("📋 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Display data info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(df))
            with col2:
                st.metric("Total Columns", len(df.columns))
            with col3:
                st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB")
            
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
    else:
        st.info("👆 Please upload an Excel file to begin")
        
        # Sample data button
        if st.button("🔄 Try with Sample Data"):
            sample_data = {
                'Product': ['A', 'B', 'C', 'A', 'B', 'C', 'A', 'B'],
                'Region': ['East', 'East', 'West', 'West', 'East', 'East', 'West', 'West'],
                'Sales': [100, 150, 200, 120, 180, 210, 110, 160],
                'Quantity': [10, 15, 20, 12, 18, 21, 11, 16]
            }
            st.session_state.df = pd.DataFrame(sample_data)
            st.session_state.tool = ExcelPythonTool()
            st.session_state.tool.df = st.session_state.df
            st.rerun()

# Tab 2: Analysis
with tab2:
    st.header("📊 Statistical Analysis")
    
    if st.session_state.df is not None:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Descriptive Statistics")
            st.dataframe(st.session_state.df.describe(), use_container_width=True)
        
        with col2:
            st.subheader("🔍 Data Types")
            dtype_df = pd.DataFrame({
                'Column': st.session_state.df.columns,
                'Type': st.session_state.df.dtypes.values,
                'Non-Null': st.session_state.df.count().values,
                'Null': st.session_state.df.isnull().sum().values
            })
            st.dataframe(dtype_df, use_container_width=True)
        
        # Missing values visualization
        st.subheader("❓ Missing Values Analysis")
        missing_data = st.session_state.df.isnull().sum()
        if missing_data.sum() > 0:
            st.bar_chart(missing_data[missing_data > 0])
        else:
            st.success("✅ No missing values found!")
    else:
        st.info("👆 Please upload a file first")

# Tab 3: Clean Data
with tab3:
    st.header("🧹 Data Cleaning")
    
    if st.session_state.df is not None and st.session_state.tool is not None:
        col1, col2 = st.columns(2)
        
        with col1:
            remove_duplicates = st.checkbox("Remove Duplicate Rows", value=True)
        
        with col2:
            fill_method = st.selectbox(
                "Handle Missing Values",
                ["mean", "median", "mode", "zero", "none"],
                index=0
            )
        
        if st.button("🧹 Clean Data", type="primary"):
            with st.spinner("Cleaning data..."):
                try:
                    original_rows = len(st.session_state.tool.df)
                    
                    if fill_method != "none":
                        st.session_state.tool.clean_data(
                            drop_duplicates=remove_duplicates,
                            fill_na_method=fill_method
                        )
                    elif remove_duplicates:
                        st.session_state.tool.df = st.session_state.tool.df.drop_duplicates()
                    
                    st.session_state.df = st.session_state.tool.df
                    new_rows = len(st.session_state.df)
                    
                    st.success(f"✅ Data cleaned! Removed {original_rows - new_rows} rows.")
                    st.dataframe(st.session_state.df.head(), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ Error cleaning data: {str(e)}")
    else:
        st.info("👆 Please upload a file first")

# Tab 4: Transform
with tab4:
    st.header("🔢 Data Transformation")
    
    if st.session_state.df is not None and st.session_state.tool is not None:
        
        # Aggregation
        st.subheader("📊 Aggregate Data")
        col1, col2 = st.columns(2)
        
        with col1:
            group_columns = st.multiselect(
                "Group By Columns",
                options=st.session_state.df.columns.tolist(),
                help="Select columns to group by"
            )
        
        with col2:
            numeric_cols = st.session_state.df.select_dtypes(include=['number']).columns.tolist()
            if numeric_cols:
                agg_column = st.selectbox("Aggregate Column", numeric_cols)
                agg_function = st.selectbox("Function", ["sum", "mean", "median", "count", "min", "max"])
        
        if group_columns and st.button("📊 Aggregate", type="primary"):
            try:
                aggregated = st.session_state.tool.aggregate_data(
                    group_columns,
                    {agg_column: agg_function}
                )
                st.success("✅ Aggregation completed!")
                st.dataframe(aggregated, use_container_width=True)
                
                # Store aggregated data
                st.session_state.aggregated = aggregated
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        
        st.markdown("---")
        
        # Download section
        st.subheader("💾 Download Processed Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 Download Current Data", type="primary"):
                # Create Excel file in memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state.df.to_excel(writer, index=False, sheet_name='Data')
                output.seek(0)
                
                st.download_button(
                    label="⬇️ Click to Download",
                    data=output,
                    file_name="processed_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col2:
            if 'aggregated' in st.session_state:
                if st.button("📥 Download Aggregated Data", type="primary"):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        st.session_state.aggregated.to_excel(writer, index=False, sheet_name='Aggregated')
                    output.seek(0)
                    
                    st.download_button(
                        label="⬇️ Click to Download",
                        data=output,
                        file_name="aggregated_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    else:
        st.info("👆 Please upload a file first")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Made with ❤️ by Jahid Hassan | 
    <a href='https://github.com/dmjahidbd/Excel-Python-Integration-Tool' target='_blank'>GitHub Repository</a>
    </p>
</div>
""", unsafe_allow_html=True)
