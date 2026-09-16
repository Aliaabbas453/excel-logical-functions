# 📊 Excel Data Analysis & Business Analytics Projects

A collection of practical **Microsoft Excel projects and assignments** focused on building strong foundations in spreadsheet-based data analysis, data manipulation, logical decision-making, lookup techniques, conditional calculations, dynamic ranges, and text processing.

This repository documents my hands-on learning journey through structured datasets and practical, business-oriented problems designed to develop skills relevant to **Business Analytics and Business Analysis**.

---

## 🚀 Skills Covered

### 🔎 Lookup & Reference Functions

- VLOOKUP
- XLOOKUP
- HLOOKUP
- INDEX
- MATCH
- INDEX + MATCH
- XMATCH
- OFFSET

### 🧠 Logical Functions

- IF
- Nested IF
- AND
- OR
- NOT
- IF + AND
- IF + OR
- IF + NOT

### 📈 Conditional Aggregation

- COUNTIF
- COUNTIFS
- SUMIF
- SUMIFS
- AVERAGEIF
- AVERAGEIFS

### 🔤 Text Functions

- CONCATENATE
- TEXTJOIN
- LEFT
- RIGHT
- MID
- FIND
- LEN
- TEXT

### 📅 Date & Formatting

- TEXT
- Date formatting
- Weekend/weekday identification
- Conditional date-based calculations
- Number and currency formatting

### 📐 Dynamic Range & Reference Techniques

- Dynamic ranges using OFFSET
- OFFSET + COUNTA
- OFFSET + SUM
- Dynamic chart ranges
- Dynamic data retrieval
- Relative position-based references

### 📌 Excel Concepts

- Relative cell referencing
- Absolute cell referencing
- Mixed cell referencing
- Formula-based data analysis
- Structured datasets
- Conditional calculations
- Data lookup and retrieval
- Dynamic references
- Business-oriented problem solving

---

# 📁 Project Structure

The repository contains multiple Excel projects, with each workbook focusing on a specific group of Excel functions and practical analytical problems.

---

## 1. 🧠 IF & Logical Functions

This project focuses on using logical formulas to evaluate conditions, classify data, and make formula-based decisions.

### Practical Applications

- Determining whether a student has passed or failed
- Assigning grades using Nested IF
- Identifying weekdays and weekends
- Calculating commission rates based on sales
- Categorizing expenses into High, Medium, and Low
- Evaluating students using IF + AND
- Determining eligibility using IF + OR
- Checking employee bonus eligibility
- Identifying products that are out of stock
- Evaluating candidate interview eligibility

### Functions Covered

- `IF`
- `Nested IF`
- `AND`
- `OR`
- `NOT`
- `IF + AND`
- `IF + OR`
- `IF + NOT`

---

## 2. 🔤 Text Functions

This project focuses on manipulating, extracting, combining, and formatting text data.

### Functions Covered

- `CONCATENATE`
- `TEXTJOIN`
- `LEFT`
- `RIGHT`
- `MID`
- `FIND`
- `LEN`
- `TEXT`

### Practical Applications

- Combining first and last names
- Extracting characters from strings
- Extracting specific portions of names
- Finding the position of characters within text
- Measuring text length
- Formatting dates
- Formatting numbers as currency
- Processing text-based datasets

---

## 3. 📈 COUNTIFS, SUMIFS & AVERAGEIFS

This project focuses on performing calculations using multiple criteria.

### Practical Applications

- Counting students who scored above a specific mark in multiple subjects
- Calculating total sales for a specific region and product category
- Calculating average salaries based on department and experience
- Counting completed orders during a particular month
- Calculating revenue for products above a specified sales value within a category

### Functions Covered

- `COUNTIF`
- `COUNTIFS`
- `SUMIF`
- `SUMIFS`
- `AVERAGEIF`
- `AVERAGEIFS`

These functions demonstrate how Excel can be used for **multi-condition business analysis and reporting**.

---

## 4. 🔎 Lookup Functions

This project focuses on retrieving information from datasets using traditional and modern Excel lookup techniques.

### Functions Covered

- `VLOOKUP`
- `XLOOKUP`
- `HLOOKUP`
- `INDEX`
- `MATCH`
- `INDEX + MATCH`

### Practical Applications

- Finding employee information using Employee ID
- Retrieving product prices
- Finding department and job information
- Matching customer records
- Looking up sales information
- Retrieving values from vertical datasets
- Retrieving values from horizontal datasets

The objective is to understand both **traditional lookup methods and modern Excel lookup techniques**.

---

## 5. 🎯 XMATCH

This project focuses on the `XMATCH` function and its applications in position-based searching and flexible matching.

### Practical Applications

- Finding the position of a specific item
- Determining student position in a sorted list
- Locating an employee within a dataset
- Finding the first value greater than a given number
- Checking whether a product exists
- Returning the position of a matched item

### Concepts Practiced

- Exact matching
- Match modes
- Searching sorted datasets
- Position-based lookup
- Dynamic existence checks
- Approximate and next-match logic

---

## 6. 📐 OFFSET & Dynamic References

This project focuses on using `OFFSET` to create flexible and dynamic references.

### Practical Applications

- Creating dynamic ranges based on user input
- Retrieving cells a specific number of rows and columns away
- Creating dynamic chart ranges
- Combining `OFFSET` with `COUNTA`
- Calculating dynamic totals using `OFFSET + SUM`
- Generating a dynamic list of the last five entries in a dataset

### Functions & Techniques

- `OFFSET`
- `COUNTA`
- `SUM`
- Dynamic ranges
- Dynamic chart references
- Relative cell movement

This project demonstrates how dynamic references can make Excel models more flexible when datasets change.

---

# 🧮 Example Formulas

## IF

```excel
=IF(E11>50,"Pass","Fail")
