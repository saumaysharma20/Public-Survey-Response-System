# 📊 Public Survey Response Analyser

### A Generic, Interactive Survey Data Analysis Dashboard Built with Python and Streamlit

**Project ID:** DA-028  
**Institution:** Invertis University, Bareilly  
**Mentor:** Mr. Divyanshu Tomar

---

## 🚀 About the Project

**Public Survey Response Analyser** is a Python-based interactive web application designed to simplify the process of survey data preparation, cleaning, analysis, visualisation and reporting.

The application allows users to upload survey datasets in **CSV and Excel (.xlsx) formats** and automatically transforms raw data into meaningful statistical information, interactive visualisations and professional reports.

The main objective of this project is to create a **generic and dynamic survey analysis system** that can work with different survey datasets without depending on fixed column names or predefined survey questions.

Instead of manually inspecting spreadsheets, calculating statistics and creating charts, users can perform these operations through a single interactive dashboard.

## 🎯 Project Objectives

- Simplify survey data analysis through an interactive dashboard.
- Support CSV and Excel file uploads.
- Automatically identify missing values, duplicate records and identifier columns.
- Provide optional data-cleaning operations.
- Dynamically classify columns based on their names and data patterns.
- Generate statistical summaries and visualisations.
- Perform group-wise and satisfaction analysis.
- Generate data-driven insights and recommendations.
- Export cleaned datasets and professional reports.

## ✨ Key Features

### 1. Data Upload and Preview
- Upload CSV and Excel files.
- Select worksheets from Excel workbooks.
- Preview original and cleaned datasets.
- Display dataset dimensions and column information.

### 2. Data Quality Analysis
- Detect missing and blank values.
- Identify duplicate records.
- Detect duplicate respondent IDs.
- Calculate an indicative data-quality score.
- Compare original and cleaned dataset quality.

### 3. Numerical Analysis
- Calculate mean, median and mode.
- Calculate standard deviation and variance.
- Display minimum, maximum and range.
- Detect potential outliers using the IQR method.
- Generate histograms, box plots and correlation heatmaps.

### 4. Response Analysis
- Analyse categorical survey responses.
- Display frequency distributions.
- Generate interactive bar charts.
- Identify commonly selected response categories.

### 5. Advanced Data Cleaning
- Remove empty rows and columns.
- Remove duplicate records.
- Trim unnecessary spaces.
- Handle missing values using available cleaning strategies.
- Remove duplicate IDs while retaining the first record.
- Compare data quality before and after cleaning.

### 6. Group-wise and Satisfaction Analysis
- Compare responses across different groups.
- Calculate group-wise averages.
- Analyse positive response percentages.
- Generate comparative charts.
- Export group analysis results to Excel.

### 7. Smart Survey Insights
- Generate rule-based findings from the uploaded dataset.
- Identify important numerical patterns.
- Highlight common response categories.
- Provide data-based recommendations.
- Generate a downloadable executive summary.

### 8. Professional Reports
- Generate formatted PDF reports using ReportLab.
- Generate multi-sheet Excel reports.
- Export cleaned datasets in CSV and Excel formats.
- Download executive summaries in TXT format.

### 9. Interactive Survey Dashboard
- Apply filters to survey data.
- Display dynamic KPIs.
- Explore response distributions.
- Generate interactive charts.
- Compare groups.
- Export filtered survey data.

## 🧠 How It Works

The application follows an eight-stage data-processing workflow:

1. **Upload:** Import a survey dataset in CSV or Excel format.
2. **Select Sheet:** Choose the required worksheet when an Excel file is uploaded.
3. **Validate:** Check the dataset for empty files and duplicate column names.
4. **Detect:** Identify blank values, identifiers, numerical fields and categorical fields.
5. **Clean:** Apply optional data-cleaning operations.
6. **Classify:** Automatically classify columns using rule-based logic.
7. **Analyse:** Perform statistical analysis, group comparisons and visualisation.
8. **Export:** Generate reports and download processed datasets.

## 🛠️ Technology Stack

| Technology | Purpose |
|---|---|
| Python 3 | Core programming language |
| Streamlit | Interactive web application |
| Pandas | Data processing and analysis |
| Plotly Express | Interactive charts and visualisations |
| OpenPyXL | Excel reading and report generation |
| ReportLab | Professional PDF report generation |
| HTML & CSS | Dashboard styling and interface design |

**Additional Python modules:** `io`, `re`, `datetime` and `xml.sax.saxutils`.

## 📈 Demonstration Results

The project presentation demonstrates the application using a sample dataset containing:

| Metric | Demonstration Result |
|---|---:|
| Survey Responses | 2,000 |
| Total Columns | 8 |
| Missing Values | 0 |
| Duplicate Records | 0 |
| Data Quality Score | 100% |
| Average Rating | 3.475 / 5 |

*Note: These are results from the demonstration dataset, not fixed values. The application analyses the dataset uploaded by the user.*

## 📂 Project Dashboard Modules

The application contains nine main modules:

1. Data Preview
2. Data Quality
3. Numerical Analysis
4. Response Analysis
5. Advanced Data Cleaning
6. Group Analysis
7. Smart Insights
8. Professional Reports
9. Interactive Dashboard

## ⚙️ Installation and Setup

### 1. Clone the Repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

### 2. Open the Project Folder

```bash
cd public-survey-response-analyser
```

### 3. Create a Virtual Environment (Recommended)

```bash
python -m venv .venv
```

Activate it on Windows:

```bash
.venv\Scripts\activate
```

### 4. Install Dependencies

```bash
pip install streamlit pandas plotly openpyxl reportlab
```

### 5. Run the Application

```bash
streamlit run app.py
```

The application will open in your browser.

## 📋 Requirements

- Python 3
- Visual Studio Code (recommended)
- Required Python libraries
- A CSV or Excel survey dataset

## ⚠️ Limitations

- Column classification is rule-based and may require manual verification for ambiguous numeric codes.
- Very large datasets may require additional memory and processing time.
- The data-quality score is an indicative project metric, not an industry-standard certification.
- Duplicate records may sometimes represent genuine responses.
- The current system does not claim sentiment analysis, topic modelling or causal analysis.

## 🔮 Future Scope

- Manual column-type selection and override.
- Advanced date-range and group filters.
- Cross-tabulation between survey variables.
- Validated sentiment and topic analysis.
- Privacy masking for sensitive survey information.
- Automated testing and performance improvements.

## 👥 Project Team

| Name | Role |
|---|---|
| Saumay Sharma | Lead Developer, Complete Coding, Dashboard Design, Data Visualisation, Testing & Project Integration |
| Vishal Gill | Research & Analysis |
| Abhi Dhariwal | Data Collection |
| Shriyansh Saxena | Documentation |
| Vishal Pathak | Project Review |
| Krishna Rohilla | UI/UX Support |
| Rohitash Kumar | Report Preparation |
| Harsh Pandey | Quality Review |
| Muskan Mourya | Presentation Design |
| Ansh Mourya | Dataset Preparation |
| Saanvi Bhardwaj | Project Coordination |
| Manish Gautam | Survey Data Validation |

## 📚 Project Documentation

- **Practical Report:** [View Project Practical Report](https://docs.google.com/document/d/1Flkcb4BkeiEab1dSf_3ZD3sj6yJl4ptD/edit?usp=sharing)
- **Project Presentation:** [View Project PPT](https://docs.google.com/presentation/d/1kfpx-Mrb8fYfToFtTfjs-CuH0PWglZGM/edit?usp=sharing)

## 🎓 Academic Information

**Project Name:** Public Survey Response Analyser  
**Project ID:** DA-028  
**Mentor:** Mr. Divyanshu Tomar  
**University:** Invertis University, Bareilly

---

## ⭐ Conclusion

Public Survey Response Analyser provides a unified platform for transforming raw survey datasets into organised, analysed and report-ready information.

By combining Python, Streamlit, Pandas and interactive visualisation tools, the project reduces repetitive manual work and makes survey data analysis more accessible, efficient and understandable.

**Built with Python | Developed at Invertis University**
