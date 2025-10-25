# 📊 Dynamic Company Financial Performance Dashboard (with Excel VBA)

### 💡 Automated Financial Data Visualization and Reporting System
This project presents a **fully Excel-based dynamic dashboard system** designed to analyze a company’s financial performance in multiple dimensions.  
The goal is to visualize sales and profitability metrics by **product, country, and period**, providing fast and interactive insights for decision-makers.

---

## ⚙️ Tools and Technologies Used
- **Microsoft Excel**
- **Pivot Tables**
- **Power Query**
- **Charts (Line, Pie, Column, Scatter)**
- **VBA (Visual Basic for Applications)**

Project Type: 📈 Data Analysis and Business Intelligence

---

## 🧩 Project Overview

### 1️⃣ Objective
To create a **dynamic dashboard** that automatically calculates and visualizes performance indicators such as profit, discount, and profit margin using sales data.

### 2️⃣ Dataset
- Period: **2013–2014**  
- Variables: Product, Country, Sales Price, Gross Sales, Net Sales, Profit, Discount, Date  
- Source: Example dataset “Retail Financial Data Sample”  
- Purpose: Educational and analytical demonstration

### 3️⃣ Data Preparation Process
Steps applied during data cleaning and transformation:
- Removed empty or invalid cells  
- Unified numeric formats (`.`, `,`)  
- Standardized date formats  
- Replaced negative or missing values with zero  
- Cleaned text fields (trimmed spaces, unified capitalization)

---

## 📊 Dashboard Structure

### 🔸 General Financial Performance Dashboard
- **KPI Cards:** Total Sales, Total Profit, Average Profit Margin  
- **Line Chart:** Sales trend over time  
- **Column Chart:** Product-based sales comparison  
- **Pie Chart:** Country-based sales share  
- **Grouped Columns:** Relationship between discount rate and profit  

### 🔸 Product & Country Analysis Dashboard
- **Column Chart:** Product-level sales  
- **Pie Chart:** Country’s share in total sales  
- **Line Chart:** Monthly country sales trends  
- **Filter (Slicer):** Dynamic filtering by product or country  

### 🔸 Profitability Analysis Dashboard
- **Column Chart:** Total profit by product  
- **Pie Chart:** Profit distribution by country  
- **Line Chart:** Monthly profit trend  
- **Scatter Plot:** Relationship between discount rate and profit margin  
- **Filters:** Country, Product, and Year selection  

---

## 🧠 VBA Automation System

### 🔹 Dashboard Switching Mechanism
Transitions between dashboards are automated using **VBA code**.  
When a user selects an option from the menu, only the corresponding dashboard becomes visible.  
This system operates on a single worksheet and uses the **Shape Visibility** method for optimized performance.

### 🔹 Code Structure (Summary)
- Dashboard switching: `Worksheet_SelectionChange`  
- Visibility management: `ShowDashboard` function  
- Error handling: `On Error Resume Next`  
- Dynamic group management (General, Product & Country, Profitability)

### 🔹 Code Advantages
- User-friendly navigation  
- Error-free and optimized visibility control  
- Modular design allowing easy addition of new dashboards  

---

## ✅ Results and Evaluation

### 🔸 Dashboard Advantages
- Access all financial KPIs from a single screen  
- Quick trend recognition via visual presentation  
- Fully dynamic and filterable layout  
- Automated logic controlled by VBA  
- Scalable and easily maintainable design  

### 🔸 General Evaluation
This project demonstrates that Excel is not only a spreadsheet tool but also a **powerful visualization and reporting platform**.  
With VBA integration, it offers a dynamic, interactive, and professional user experience — ideal for analysts and managers alike.

---

## 👨‍💻 Author
**Ahmet Dokazoğlu**  
📍 Ankara, Türkiye  
🔗 [GitHub Profile](https://github.com/AhmetDokazoglu)  
🔗 [LinkedIn Profile](https://www.linkedin.com/in/ahmet-dokazo%C4%9Flu-9660b2346/)

---

## 📎 Additional Documents
📄 [Download the Full Project Report (Word Version)](https://github.com/AhmetDokazoglu/Dynamic-Company-Financial-Performance-Dashboard--with-Excel-VBA-/raw/refs/heads/main/Dynamic%20Company%20Financial%20Performance%20Dashboard%20(with%20Excel%20VBA)(TR).docx)
