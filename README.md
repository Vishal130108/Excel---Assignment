 Product Dataset Analysis using Excel

## 📌 Project Overview
This project involves performing data analysis on a Product Dataset using Microsoft Excel. The objective is to explore, analyze, and transform raw product data using built-in Excel functions to derive meaningful business insights.

The analysis includes statistical calculations, conditional logic, and text-based data extraction to enhance data usability and interpretation.

---

## 📂 Dataset Description

**Dataset Name:** Product Dataset  

**Attributes:**
- Product ID  
- Product Name  
- Brand Name  
- Quantity  
- Category  
- Price  

---

## 🎯 Objectives
- Perform basic data exploration using Excel functions  
- Apply logical and conditional functions for data classification  
- Extract meaningful insights using aggregation functions  
- Transform raw data using text functions  

---

## 🧮 Tasks Performed

---

## 1. 📊 Basic Data Exploration

The following calculations were performed:

- Total price of all products using **SUM**
- Total number of products using **COUNTA**
- Average price of products using **AVERAGE**

---

## 2. 📉 Minimum and Maximum Values

- Minimum product price identified using **MIN**
- Maximum product price identified using **MAX**

---

## 3. 🧠 Logical Function – IF

A new column **Price Range** was created using the IF function:

### Classification Rules:
- If Price ≥ 500 → **High Price**
- If Price < 500 → **Standard Price**

### Formula:
```excel
=IF(Price>=500,"High Price","Standard Price")
4. 🔍 Conditional Functions – SUMIF & COUNTIF
✔ Total Price of Electronics Category
=SUMIF(CategoryRange,"Electronics",PriceRange)
✔ Count of Products with Price < $100
=COUNTIF(PriceRange,"<100")
5. 🧩 Text Functions – Data Extraction

Product ID Format Example: 28-JAN-US

✔ Day Column (First 2 Characters)
=LEFT(ProductID,2)
✔ Month Column (4th to 6th Characters)
=MID(ProductID,4,3)
✔ Country Code Column (Last 2 Characters)
=RIGHT(ProductID,2)
📊 Key Outcomes
Products classified based on pricing structure
Category-wise financial insights generated
Low-price product distribution analyzed
Product ID successfully structured into meaningful components
Improved data readability through transformation
🛠 Tools Used
Microsoft Excel
🚀 Skills Demonstrated
Data Analysis
Data Cleaning & Transformation
Logical Functions (IF)
Conditional Functions (SUMIF, COUNTIF)
Text Functions (LEFT, MID, RIGHT)
Business Data Interpretation
👨‍💻 Author

Vishal Kannan
Aspiring Data Analyst
