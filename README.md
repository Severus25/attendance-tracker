# SPRAC Attendance | Workforce Analytics Portal

**A centralized monthly employee attendance reporting and workforce analytics application developed for SPRAC Services.**

SPRAC Attendance transforms monthly employee attendance PDF reports into structured, management-ready Excel workbooks. It automates attendance analysis, calculates employee-wise monthly averages, and consolidates daily attendance records into a single, easy-to-understand report.

---

## Overview

Managing monthly attendance records across multiple employees can be time-consuming and prone to manual calculation errors.

SPRAC Attendance simplifies this process by allowing HR and operations teams to upload multiple employee attendance PDFs simultaneously and automatically generate a consolidated Excel report.

The application is designed with a clean, intuitive interface so that users can operate it without any technical knowledge.

## Key Features

### 1. Bulk PDF Upload
- Upload multiple employee attendance PDF reports in a single batch.
- Supports processing approximately 10–20 employee reports per batch.
- Accepts PDF attendance reports.
- Designed for reports belonging to the same calendar month.

### 2. Automated Attendance Analysis

For every employee, the application calculates:

| Metric | Description |
|---|---|
| Average In Time | Average daily employee arrival time |
| Average Out Time | Average daily employee departure time |
| Average Stay Time | Average time spent in the office per recorded working day |
| Daily Attendance | Individual daily attendance and stay duration |

### 3. Consolidated Excel Reporting

The generated Excel workbook organizes employee information in a structured format.

**Report structure:**

1. Employee ID
2. Average In Time
3. Average Out Time
4. Average Stay Time
5. Day-wise attendance records for the entire month

The workbook is designed for easy review, sharing, and monthly HR reporting.

### 4. User-Friendly Interface
- Professional dashboard designed for HR teams.
- Light and dark appearance options.
- Simple instructions for non-technical users.
- Clear upload and processing workflow.
- Downloadable Excel reports.

---

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application logic |
| Streamlit | Web application and user interface |
| PDFPlumber | PDF text extraction |
| OpenPyXL | Excel workbook generation |
| Regular Expressions | Attendance data parsing |
| Python Calendar & Datetime | Date and time calculations |

---

## Project Structure

```text
SPRAC-Attendance/
│
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## Installation and Local Setup

### Prerequisites

- Python 3.9 or later
- pip package manager
- Git (optional, for cloning the repository)

### Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR-USERNAME/YOUR-REPOSITORY.git
```

Navigate into the project directory:

```bash
cd YOUR-REPOSITORY
```

### Step 2: Create a Virtual Environment

```bash
python -m venv venv
```

Activate it:

**Windows:**

```bash
venv\Scripts\activate
```

**macOS / Linux:**

```bash
source venv/bin/activate
```

### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4: Launch the Application

```bash
streamlit run app.py
```

The application will open in your browser, generally at:

```text
http://localhost:8501
```

---

## How to Use the Application

1. Open the SPRAC Attendance portal.
2. Select your preferred light or dark appearance.
3. Click the PDF upload area and select the employee attendance reports.
4. Ensure all uploaded reports belong to the same month.
5. Review the processed employee records and any import warnings.
6. Download the consolidated Excel workbook.
7. Share the generated workbook with the concerned HR or management team.

No programming knowledge is required to use the application.

---

## Input Requirements

- Original attendance-system PDF reports should be used.
- Each employee should have an individual attendance report.
- All reports in one batch should belong to the same month.
- If the employee ID is unavailable in the PDF, the filename may be used as a fallback, where supported.
- PDFs should contain extractable text. Scanned image-only PDFs may require OCR support.

---

## Output

The application generates an Excel workbook containing employee-wise monthly attendance analytics and day-wise attendance records.

The workbook is intended to support:
- Monthly HR attendance reviews.
- Employee attendance analysis.
- Workforce reporting.
- Management-level attendance summaries.

---

## Important Notes

- The application is designed for monthly attendance processing.
- Uploaded reports are processed by the application; this project does not provide a permanent employee database.
- Processing accuracy depends on the structure and readability of the source attendance PDFs.
- Always verify exceptional attendance records before using reports for payroll or other official decisions.
- Appropriate access controls should be implemented before processing confidential employee information in a publicly accessible deployment.

---

## Deployment

This application is built using Streamlit.

For deployment instructions, refer to the official [Streamlit Community Cloud documentation](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).

A conventional Streamlit application should be hosted on a platform that supports running Streamlit's application server.

---

## Developed For

**SPRAC Services**

Staffing | Projects | Research | Advisory | Consulting

Website: https://spracservices.com/

---

*SPRAC Attendance — Workforce Analytics Portal*  
*Internal HR Operations | Monthly Attendance Reporting*
