# Auto for Jira

A script meant to fully automate processes currently performed manually.

## Project Overview

This project comprises multiple Python scripts, each designed for a specific purpose. Below is an overview of these scripts and their functionalities:

### 1. Script 1: work_ratio.py

- **Description**: Generates pie charts illustrating work ratios based on various parsers. Outputs an Excel file containing the data used for the charts.

- **Dependencies**:
  - `worker-names.csv` (static CSV file, updated only when team member names change).
  - `v1-yearly-jira.csv` (dynamic CSV file, updated for each run with yearly work ratio data from Jira).

### 2. Script 2: another_script.py

- **Description**: Generates an Excel file displaying monthly planned, actual, and usage differences for BI, Algo, and Dev teams. Also includes yearly totals.

- **Dependencies**: 
  - `worker-names.csv` (static CSV file, updated only when team member names change).
  - `budget-naming.xlsx` (static Excel file, updated when budget names change).
  - `yearly-company-budget.xlsx` (static Excel file, updated when yearly budgets change).
  - `jira-missions-yearly2023.csv` (dynamic CSV file, updated for each run with yearly data from Jira).

## Usage

Provide detailed instructions on how to run each script, including command-line arguments and configuration settings. Include example commands where applicable.

## Installation

To set up the environment for these scripts, follow these steps:

1. Create and activate a virtual environment (e.g., myenv).
2. Install the required libraries using `pip` or another package manager:
   - python>=3.10.7
   - pandas
   - calendar
   - openpyxl
   - openpyxl.styles
   - matplotlib.pyplot
   - plotly.express
   - numpy
   - argparse
   - datetime

## Configuration

Before running the scripts, ensure the following:

- There are no Hebrew words in any of the Excel or CSV files, as pandas uses UTF-8 encoding.
