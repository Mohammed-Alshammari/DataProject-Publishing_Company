# 📊 Sales Analysis and Reporting for a Publishing Company Using Excel



## 📌 Project Overview
This project focuses on analyzing sales data for West Cal Publishing(fictional entity) using Microsoft Excel. The dataset contains sales records from different regions, and the objective is to clean, structure, and visualize the data effectively for senior management. The project is an honors assignment for the Microsoft Excel course.

## 📂 Data Source
- **Sample File - Original.xlsx**: Contains the raw sales data.
- **New Staff.csv**: Includes details of new employees to be integrated into the dataset.
 
 Both were obtaind from the microsoft excel course.


## 🛠 Tools
- Microsoft Excel


## 🔍 Data Cleaning & Preparation
1. Renamed the original worksheet to All Sales and sorted the data by Month and Sales Area to orgnize it.
2. Added new columns:
   - **Targets** (Fixed value of 15,000 across all rows)
   - **Commission** (Calculated using an IF function based on sales performance)
3. Created separate sheets (North, South, East, West) by copying the All Sales sheet and then filtiring based on Sales Area.
4. Applied Conditional Formatting to highlight top 5 sales values in each region.

## 📊 Exploratory Data Analysis
- Applied Conditional Formatting on each area sheet to highlight top 5 sales values in each region.
- Made the the name of the sales people in each region sheet as headings and calculated their total sales
  ![image](https://github.com/user-attachments/assets/35b3db0c-1bc1-4c88-b05d-dfde4240e359)

  ![image](https://github.com/user-attachments/assets/db16f03b-a460-482b-9b3a-2cf169939782)

  ![image](https://github.com/user-attachments/assets/24119441-47fc-47c6-8e31-18eaa1a63fdb)

  ![image](https://github.com/user-attachments/assets/6bb910b0-5b09-4db0-a734-6998d5739d1c)




  



## 📈 Data Analysis & Visualization
- Copied the All Sales sheet, renamed it Copy of All Sales, and cleared formatting from the titles. Converted the data into a table with the Total Row enabled, adding totals in the Sales Amount and Commission columns. Named the table Sales_Data and applied a style.
- Added the Over/Under column with a formula that calculates the diffrences between Sales Amount and Target using table references, formatting the results as currency with "negative values in red, "Postive" values in green.
---
![image](https://github.com/user-attachments/assets/5674e57e-b5d8-43a5-a265-22fa2f3e17ed)

---
- Created a sheet called Chart and added a summary block with area totals for sales and commission. Used the Total Row of the Table to generate totals.
- Created a Column Chart from the data that shows total sales and commissions per region.
---
  ![image](https://github.com/user-attachments/assets/26df0ed5-e292-45c4-9dd4-d47f5007bc4b)
  
 --- 
  
  
- Inserted a PivotTable in a new worksheet named Sales Analysis, showing sales totals by month formatted as currency.
- Added a second set of totals as a percentage of the grand total and enabled filtering by payment type.
- Added Sales Area and Employee slicers.
---
![image](https://github.com/user-attachments/assets/4af680f1-c98d-4c1c-80b1-769aae8695cd)

---


- Imported data from the New Staff CSV file via the Data Ribbon, broke the link, removed connections, and closed query windows.
- Added blank columns, transforming staff names using text functions to match the sample. Extracted area initials from the Payroll ID column and hid all columns except First Name, Last Name, Area and Payroll Code.
---
![image](https://github.com/user-attachments/assets/c1525bdc-256a-416c-9414-88fa32cec35a)

---

 
- Created a Cover Sheet.
- Inserted hyperlinks for navigation.
---
![image](https://github.com/user-attachments/assets/3ce0e938-b747-4a95-b197-a72a1e1012b0)

---


   
## 🔬 Results & Findings
- The sales performance varied significantly across regions.
- Conditional formatting made it easy to spot top-performing salespeople.
- The structured pivot table provided an intuitive breakdown of revenue distribution.
- The visualizations simplified the communication of insights to senior management.

## 💡 Recommendations
- Focus on underperforming regions by analyzing sales trends and setting realistic targets.
- Adjust commission rates based on performance to incentivize higher sales.
- Implement similar structured reporting templates for future sales analysis.

## 📑 Additional Features
- **Copy of All Sales**: Reformatted data as a structured table (Sales_Data) with auto-calculated totals.
- **New Staff Sheet**: Processed and formatted new employee data using text functions.
- **Cover Page**: Designed with hyperlinks for seamless navigation across sheets.

This project demonstrates my ability to manage, analyze, and present data effectively using Excel, reinforcing my skills in data-driven decision-making. 🚀
