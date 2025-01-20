# Automated-Report-Generation

This project generates a detailed sales report using data from a CSV file. The report includes several visualizations, such as a bar graph and a pie chart, as well as structured data tables and conclusions. The report is generated in PDF format and can be used for analysis and decision-making.

## Features

- Reads sales data from a CSV file (`salesreport.csv`).
- Generates visualizations like:
  - A **Discount vs. Total Sales Bar Chart**.
  - A **Sales Trends Pie Chart** based on product-wise sales.
- Creates formatted tables for:
  - Product-wise performance.
  - Region-wise performance.
  - Salesperson-wise performance.
- Exports the final report as a PDF document with tables, charts, and a narrative analysis.

## Prerequisites

Before running the code, ensure that you have the following libraries installed:

- `pandas` - for handling CSV data and performing data aggregation.
- `matplotlib` - for creating charts and visualizations.
- `reportlab` - for generating PDFs.
- `csv` - for reading CSV files.

## PDF Generation:

- A structured sales report is created using the reportlab library.
- The report includes:
  - An introductory section with the objectives.
  - Key data tables (product-wise, region-wise, salesperson-wise).
  - Visualizations embedded in the report.
  - Conclusions derived from the data.

### Installing the required libraries:

You can install the required libraries using `pip`:

```bash
pip install pandas matplotlib reportlab
