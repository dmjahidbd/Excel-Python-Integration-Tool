# Excel-Python Integration Tool

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Active-success)

## 🚀 [**Try Live Web App**](https://excel-python-integration-tool-b4xfv3dxnz4cufkjpe4foo.streamlit.app/) 🌐

> **No installation needed!** Upload Excel files and analyze data directly in your browser.

A comprehensive Python tool for seamless Excel integration - automate data analysis, apply machine learning models, and perform advanced analytics directly with Excel files.

## Features

- **Excel Data Loading**: Read Excel files into pandas DataFrames with ease
- **Data Cleaning**: Remove duplicates, handle missing values intelligently
- **Custom Formulas**: Apply Python-based formulas to create calculated columns
- **Data Filtering**: Filter data based on complex conditions
- **Aggregation**: Group and aggregate data with multiple functions
- **Pivot Tables**: Create dynamic pivot table analysis
- **Styled Export**: Export data to Excel with professional formatting
- **Statistical Analysis**: Perform comprehensive data analysis
- **ML Ready**: Integrate with scikit-learn for predictions

## Installation

### Basic Installation

```bash
# Clone the repository
git clone https://github.com/dmjahidbd/Excel-Python-Integration-Tool.git
cd Excel-Python-Integration-Tool

# Install core dependencies
pip install pandas numpy openpyxl
```

### Full Installation (with all features)

```bash
pip install -r requirements.txt
```

## Quick Start

### Basic Usage

```python
from excel_python_tool import ExcelPythonTool

# Initialize the tool
tool = ExcelPythonTool('data.xlsx')

# Load data
tool.load_excel()

# Perform basic analysis
analysis = tool.basic_analysis()
print(analysis)

# Clean data
tool.clean_data(drop_duplicates=True, fill_na_method='mean')

# Export results
tool.export_to_excel('output.xlsx', styled=True)
```

### Advanced Usage

```python
# Apply custom formula
tool.apply_formula('Total_Revenue', 
                   lambda row: row['Price'] * row['Quantity'])

# Filter data
filtered = tool.filter_data(tool.df['Sales'] > 1000)

# Aggregate data
aggregated = tool.aggregate_data(
    ['Product', 'Region'],
    {'Sales': 'sum', 'Quantity': 'mean'}
)

# Create pivot table
pivot = tool.pivot_analysis(
    index='Product',
    columns='Region',
    values='Sales',
    aggfunc='sum'
)
```

## Demo

Run the included demo to see the tool in action:

```bash
python excel_python_tool.py
```

This will:
1. Create sample data
2. Load and analyze it
3. Apply formulas and transformations
4. Generate styled Excel reports

## Use Cases

### Business Analytics
- **Sales Analysis**: Aggregate sales data by product, region, or time period
- **Financial Reporting**: Generate automated financial summaries
- **Inventory Management**: Track stock levels and calculate reorder points

### Data Science
- **Data Preprocessing**: Clean and prepare Excel data for ML models
- **Feature Engineering**: Create calculated fields for predictive models
- **Results Export**: Save model predictions back to Excel

### Automation
- **Report Generation**: Automate weekly/monthly report creation
- **Data Validation**: Check data quality and flag anomalies
- **Batch Processing**: Process multiple Excel files at once

## Project Structure

```
Excel-Python-Integration-Tool/
├── excel_python_tool.py    # Main tool implementation
├── requirements.txt         # Project dependencies
├── README.md               # This file
├── .gitignore             # Git ignore file
└── examples/              # Usage examples (coming soon)
```

## Requirements

### Core Dependencies
- pandas >= 2.0.0
- numpy >= 1.24.0
- openpyxl >= 3.1.0

### Optional Dependencies
- xlwings >= 0.30.0 (for real-time Excel integration)
- scikit-learn >= 1.3.0 (for machine learning)
- matplotlib >= 3.7.0 (for visualizations)

See `requirements.txt` for the complete list.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License.

## Author

**D M Jahid Hasan**
- GitHub: [@dmjahidbd](https://github.com/dmjahidbd)

## Acknowledgments

- Built with pandas for data manipulation
- Uses openpyxl for Excel file handling
- Inspired by the need for better Excel-Python workflow automation

## Support

If you find this tool useful, please consider giving it a star ⭐ on GitHub!

---

**Note**: This tool is designed for data analysis and automation. Always validate your results and test with sample data first.
