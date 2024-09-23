# 🚀 Python GSheet Automation

### Objective:
**Automate courier performance analysis using Google Sheets** by importing data, calculating KPIs, and generating automated reports to provide insights into courier efficiency and effectiveness.

---

## 📂 Dataset Description:
1. **CourierData**:
   - `CourierID`: Unique ID for each courier.
   - `Name`: Courier's name.
   - `Region`: Courier’s operating region.
   - `Deliveries`: Number of deliveries per day.
   - `Date`: Date of deliveries.
   - `PackageWeight`: Weight of delivered packages.
   - `Month`: Derived from the date (Month and Year).

2. **Performance Data**:
   - `CourierID`: Unique ID for each courier.
   - `TotalDeliveries`: Total deliveries made.
   - `AverageWeight`: Average package weight.
   - `BestMonth`: The month with the highest deliveries.

---

## 🛠️ Key Features:

- **Automated Data Import**: Fetch courier data from an external Google Sheet.
- **KPI Calculation**: Automatically calculate `TotalDeliveries`, `AverageWeight`, and determine `BestMonth`.
- **Macros**: Automate data refresh and update pivot tables.
- **Pivot Tables**: Summarize courier performance data and insights.

---

## 💡 Tasks Breakdown:

### 1. **Data Importation and Cleaning**:
- Import **CourierData** from another Google Sheet.
- Ensure correct data formatting and calculate the **Month** column from the **Date**.

### 2. **KPI Calculations**:
- Use `SUMIF` and `AVERAGEIF` for key metrics like **TotalDeliveries** and **AverageWeight**.
- Build a pivot table to identify the **BestMonth** for each courier.

### 3. **Macro Automation**:
- Record and implement a macro for automated data refresh (includes updating `IMPORTRANGE` data and refreshing the pivot table).

---

## 💻 Code Examples:

### **Data Import Formula**:
```python
=IMPORTRANGE("URL_of_the_Google_Sheet", "CourierData!A1:G")



