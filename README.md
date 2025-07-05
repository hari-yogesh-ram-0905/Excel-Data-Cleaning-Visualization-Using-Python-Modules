
# Excel Data Cleaning & Visualization Using Python Modules

---

## Table of Contents

1. [Overview](#overview)  
2. [Motivation](#motivation)  
3. [Goals](#goals)  
4. [Tech Stack](#tech-stack)  
5. [Features](#features)  
6. [Data Description](#data-description)  
7. [Installation](#installation)  
8. [Usage](#usage)  
9. [Results](#results)  
10. [Future Work](#future-work)  
11. [Contact](#contact)

---

## 1. Overview

This project is an end-to-end data engineering and visualization tool built with Python. It allows users to load multiple Excel files, merge, clean, and store the cleaned data into both a new Excel file and a MySQL database. Finally, it provides an interactive GUI to perform various visualizations on the cleaned data.

---

## 2. Motivation

Dealing with scattered, uncleaned Excel files is a routine challenge in many organizations. Automating data cleaning and visualization reduces errors, saves time, and builds a reusable and scalable workflow for data preparation and insights.

---

## 3. Goals

- Automate the cleaning of multiple Excel datasets  
- Provide a GUI to make the workflow user-friendly  
- Save cleaned data in both Excel and a relational database  
- Enable rapid visualization for exploratory data analysis  
- Demonstrate best practices in Python-based data engineering

---

## 4. Tech Stack

- **Programming Language**: Python  
- **Data Processing**: pandas, numpy  
- **Visualization**: matplotlib, seaborn  
- **GUI**: tkinter  
- **Database**: MySQL  
- **Others**: glob, os

---

## 5. Features

- Load and merge multiple Excel (`.xlsx`) files from a folder  
- Clean the merged data by removing duplicates and null values  
- Save the cleaned data into a new Excel file  
- Dynamically create a MySQL table from cleaned data and insert rows  
- Visualize the cleaned data using:  
  - Line plot  
  - Bar chart  
  - Pie chart  
  - Scatter plot  
  - Histogram  
  - Heatmap  
- User-friendly multi-tabbed GUI with tkinter  
- Scrollable previews of cleaned data

---

## 6. Data Description

The sample data used includes five Excel files named:

- `sales_data_1.xlsx`  
- `sales_data_2.xlsx`  
- `sales_data_3.xlsx`  
- `sales_data_4.xlsx`  
- `sales_data_5.xlsx`

These files contain sales transactions data, with fields like `OrderID`, `Customer`, `Product`, `Quantity`, `Price`, etc.

After cleaning, the merged file is saved as `cleaned_data.xlsx`.

---

## 7. Installation

1. Clone or download this project folder  
2. Make sure you have Python 3.x installed  
3. Install dependencies:

```bash
pip install pandas numpy matplotlib seaborn mysql-connector-python
```

4. Set up MySQL locally and create a database called `Python_Project`

5. Place your input Excel files in the `dataset/` folder

---

## 8. Usage

- Run the Python script:

```bash
python your_script.py
```

- The GUI will appear with multiple tabs:
  - **About**: project description, steps involved in the Project 
  - **Data Cleaning**: select folder, clean data, preview, save to Excel and MySQL  
  - **Visualizations**: upload cleaned Excel and visualize the particular plots
  - **Conclusion**: final results

Follow the step-by-step instructions on the GUI to clean and explore your data.

---

## 9. Results

After running the workflow:

- A cleaned Excel file `cleaned_data.xlsx` is created  
- Data is stored in a MySQL table dynamically based on column types  
- Interactive visualizations are available for better understanding of the cleaned dataset  
- The GUI demonstrates a smooth, end-to-end data engineering pipeline from ingestion to insights

---

## 10. Future Work

- Add advanced statistical analysis  
- Handle categorical encoding  
- Export interactive visualizations as images  
- Integrate with web frameworks (e.g., Streamlit, Flask)  
- Provide user authentication and role-based access for larger teams

---

## 11. Contact

For any questions, feedback, or collaboration opportunities, please contact:  

**Hari Yogesh Ram B**  
Email: hariyogeshram882@gmail.com
 
---
