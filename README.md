# Excel-Minor-Project
This repository is my first Excel Minor Project.

# 📊 Excel Sales Operations Analytics Project

![Excel](https://img.shields.io/badge/Excel-Advanced-green) ![Data Analysis](https://img.shields.io/badge/Data_Analytics-100%25-brightgreen) ![Assessment](https://img.shields.io/badge/Assessment-Interview--grade-blue)

## 🎯 Project Overview

A comprehensive sales operations analytics assessment that simulates real-world data challenges faced by analysts at fast-growing multi-category retailers. This project demonstrates advanced Excel skills across lookup functions, data cleaning, conditional aggregation, pivot tables, and advanced analytics.

## 📈 Project Scope

- **2,000+ orders** analyzed across 6 Indian cities
- **5 data sheets**: Orders, Customers, Products, Sales_Reps, Targets
- **30 questions** across 5 sections (100 marks total)
- **Real-world messy data**: duplicates, missing references, inconsistent case, invalid records

## 🛠️ Skills Demonstrated

### Q1: Lookup Functions (25 marks)
- ✅ VLOOKUP for basic data retrieval
- ✅ INDEX-MATCH for flexible lookups
- ✅ Two-way lookups (row + column)
- ✅ Chain lookups across multiple sheets
- ✅ Approximate-match lookups
- ✅ Self-joins for hierarchical data

### Q2: Data Cleaning & Quality Diagnostics (20 marks)
- ✅ Duplicate detection and counting
- ✅ Whitespace identification (leading/trailing)
- ✅ Case standardization
- ✅ Referential integrity validation
- ✅ Discontinued SKU identification
- ✅ Suspicious pattern detection (promo abuse)

### Q3: Conditional Aggregation (25 marks)
- ✅ SUMIFS with multiple criteria
- ✅ COUNTIFS with complex conditions
- ✅ AVERAGEIFS with cross-sheet data
- ✅ Top-N analysis with LARGE functions
- ✅ Performance tier classification
- ✅ Multi-condition revenue analysis

### Q4: Pivot Tables (20 marks)
- ✅ 2D pivot table construction
- ✅ Cross-sheet data integration
- ✅ Percentage of row total analysis
- ✅ Date grouping by quarter
- ✅ Calculated fields for profit margins

### Q5: Advanced Analytics (10 marks)
- ✅ RANK functions for performance analysis
- ✅ Running totals and cumulative calculations
- ✅ TEXTJOIN for data concatenation
- ✅ LARGE + INDEX/MATCH combinations
- ✅ Multi-condition top-N analysis

## 🚀 Key Findings

### Data Quality Issues Identified
- **20 duplicate order records** detected
- **15 customer names** with whitespace issues
- **30 broken foreign key references** (orders pointing to non-existent customers)
- **174 orders** referencing discontinued products
- **58 suspicious patterns**: Cancelled orders with >20% discount (potential promo abuse)

### Business Insights
- **Total Revenue Analyzed**: ₹18M+ (Jan-Aug 2024)
- **Top Performing Region**: South (₹1.3M in Electronics)
- **Average Profit Margin**: 44.88% (Delhi shipments)
- **Q4 2024 Revenue**: ₹6.46M (Oct-Dec period)
- **Top Sales Rep**: Neha Desai (North region)

## 📊 Sample Excel Formulas Used

```excel
# Advanced Lookup with INDEX-MATCH
=INDEX(Sales_Reps!B:B,MATCH(VLOOKUP("ORD-00987",Orders!A:K,5,FALSE),Sales_Reps!A:A,0))

# Two-way Lookup for Targets
=INDEX(Targets!A:M,MATCH("R012",Targets!A:A,0),MATCH("Mar-24",Targets!1:1,0))

# Data Quality - Duplicate Detection
=SUMPRODUCT(--(COUNTIF(Orders!A2:A2001,Orders!A2:A2001)>1))

# Conditional Aggregation
=SUMIFS(Orders!J:J,Orders!B:B,">=2024-07-01",Orders!B:B,"<=2024-09-30",Orders!K:K,"Mumbai")

# Multi-condition Count
=COUNTIFS(Orders!K2:K2001,"Cancelled",Orders!I2:I2001,">20")
```

## 🛡️ Assessment Rules Followed

✅ Formulas preferred over hardcoded values  
✅ No modifications to original data sheets  
✅ Helper columns used where required  
✅ All answers in yellow cells as specified  
✅ Proper formatting (Currency, Percentage, Text)  

## 🎓 Learning Outcomes

This project demonstrates:
- **Advanced Excel mastery** beyond basic spreadsheet skills
- **Data cleaning expertise** handling real-world messy data
- **Business acumen** translating data into actionable insights
- **Problem-solving skills** tackling complex multi-step analyses
- **Attention to detail** identifying subtle data quality issues

## 🔧 Technical Implementation

- **Primary Tool**: Microsoft Excel with advanced functions
- **Validation**: Python/pandas for answer verification
- **Formula Types**: VLOOKUP, INDEX-MATCH, SUMIFS, COUNTIFS, Pivot Tables
- **Data Volume**: 2,000 orders × 11 columns, 300 customers, 60 products, 30 sales reps

## 🤝 Contributing

This is a completed assessment project. Feel free to use it as a reference for learning advanced Excel techniques or as a template for similar analytics projects.

## 📄 License

This project is for educational purposes. Please credit the original assessment source if used for learning or demonstration.

## 📞 Contact

For questions about the methodologies or formulas used in this project, feel free to reach out or open an issue in the repository.

---

**Project Completed**: August 2025  
**Assessment Type**: Interview-Grade Data Analytics  
**Difficulty Level**: Advanced  
**Time Required**: ~4-6 hours

#Excel #DataAnalytics #BusinessIntelligence #SalesAnalytics #DataCleaning #PivotTables #AdvancedExcel

