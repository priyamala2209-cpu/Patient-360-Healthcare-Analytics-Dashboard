# 🏥 AETNA HEALTHCARE 360 – Integrated Healthcare Analytics and Revenue Cycle Dashboard

> Excel Analysis + Power BI DAX + Interactive Dashboard Documentation
> Created By: Priyadharshini Naresh D | Program In AI Driven Data Analysts | Batch: FNB16

![Power BI](https://img.shields.io/badge/Power%20BI-DAX-yellow)
![Excel](https://img.shields.io/badge/Excel-Advanced%20Analysis-green)
![Healthcare](https://img.shields.io/badge/Domain-Healthcare%20Analytics-blue)
![Status](https://img.shields.io/badge/Project-Completed-success)

---

Excel Drive Link: https://drive.google.com/drive/u/0/folders/1dXcB3Hq98RTeqdkUr56PM9fttvO8l5pr

## 📌 Project Overview

This project analyses an integrated healthcare dataset using **Excel and Power BI**. The Excel workbook contains cleaned/master tables, calculated fields and PivotTable-based summaries, while the PBIX file contains a multi-page interactive Power BI report with DAX measures, cards, charts, tables and slicers.

**Health Track 360** is a data-driven health analytics project focused on transforming raw health data into clear, actionable insights. The process involved data cleaning and transformation in Excel using advanced functions, descriptive statistics, and pivot tables. The cleaned data was then modelled in Power BI with table relationships and visualized through an interactive dashboard.

This end-to-end workflow helps healthcare stakeholders to track key health metrics, identify trends and risk areas, and make informed, data-backed decisions.

The project demonstrates how patient, clinical, operational and financial data can be consolidated into a single **decision-support framework** supporting descriptive, diagnostic, predictive and prescriptive analysis.

---

## 🎯 Objectives

The main objective is to convert patient, encounter, clinical, procedure, prescription, provider, vitals and insurance-claim data into actionable insights.

1. **Patient Demographics Analysis:** Analyse patient demographics and routine-care patterns to understand the population served.
2. **Encounter & Revenue Analysis:** Evaluate encounter volume and charge revenue by encounter type, CPT/service and facility.
3. **Clinical Insights:** Identify major diagnoses and clinical patterns contributing to healthcare activity and charges.
4. **Provider Performance:** Measure provider and specialty performance to identify high revenue providers and staffing patterns.
5. **Pharmacy Performance:** Assess prescription volume, medication revenue and pharmacy performance.
6. **Revenue Cycle Management:** Evaluate insurance billing, collections and outstanding balances to support revenue-cycle decisions.
7. **Population Health:** Monitor vital-sign and risk indicators for population-health oriented analysis.

---

## 🗂️ Data Sources

- **Source:** Healthcare dataset obtained from ExcelX - The Ultimate Healthcare Dataset Generator
- **Timeline:** Data covers period 2022-2026
- **Domain:** Healthcare Analytics / Public Health / Hospital Management
- **Link:** Excelx.com
- **Volume:**
    - 1,000 Patients Master Records
    - 5,491 Encounters
    - 5,073 Insurance Claims
    - 2,656 Prescriptions
    - 5,660 Procedure Records
    - 5,614 Vitals Records

### Key Attributes:

| Table | Key Columns | Description |
| :--- | :--- | :--- |
| **Patients_Master** | PatientID, DOB, Age, Gender, BloodType | Demographic and clinical identity |
| **Encounters_Log** | EncounterID, Date, Type, ProviderID, ChargeAmount | Core fact table for visit & revenue |
| **Charges / Insurance_Claims** | ClaimID, AmountBilled, AmountPaid, Remaining | Revenue Cycle - billing, collection, AR |
| **Diagnoses** | ICD Code, Description | Population health & risk scoring |
| **Procedures** | CPT Code, Minutes, PatientType | Service volume and procedure cost |
| **Prescriptions** | Medication Name, Pharmacy Name, Charges | Pharmacy performance |
| **Vitals_Log** | SystolicBP, DiastolicBP, HeartRate, Temperature | Risk indicators |

---

## 🛠️ Tools & Technologies

- **Microsoft Excel:** Data cleaning, helper columns, categorisation, PivotTables, XLOOKUP, descriptive statistics (Mean, Median, Mode, SD), summaries and validation checks.
- **Power BI Desktop:** Data modelling, DAX measures, calculated columns, KPI cards, charts, slicers, drill-down and consolidated reporting.
- **Power Query:** Data transformation, type conversion, and normalization.
- **DAX (Data Analysis Expressions):** For dynamic business logic and KPIs.

---

## 🔧 Data Preprocessing

All preprocessing was done in **Excel & Power Query Editor**:

- **Data Cleaning:** Removed/ignored blank analysis rows by validating primary-key fields (PatientID, EncounterID, ProviderID, FacilityID, ChargeID, ClaimID) before analysis.
- **Standardization:** Standardized identifier fields and text fields using TRIM/CLEAN. Converted text dates to proper Date format.
- **Derived Categories Created:** Age Group, Charge Tier (Low/Mid/High), Service Type, Claim-Size Category (Small/Medium/Large), Medication Category (Antibiotic, Diabetes, BP, Pain Relief), Clinical Risk Categories, BP Category, Pharmacy Chain/Type.
- **Pivot Validation:** Used PivotTables for demographic, revenue, diagnosis, procedure, pharmacy, insurance and provider summaries. Validated totals across linked summaries (e.g., 5,614 charge records).
- **Model Structure:** Built logical fact/dimension structure:
    - **Facts:** Encounters / Charges / Claims / Prescriptions / Procedures / Vitals / Diagnoses
    - **Dimensions:** Patients, Providers, Facilities, Medications, Pharmacies

---

## 📐 Data Modeling & DAX

**Model Type:** Star Schema with 13 tables in PBIX
**Tables:** `Charges, Diagnoses, Encounters_Log, Facilities_Master, Insurance_Claims, Medications_Master, Patient_Contacts, Patients_Master, Pharmacies_Master, Prescriptions, Procedures, Providers_Master, Vitals_Log`
**Pages:** 10 Analytical Pages

### Key DAX Measures

**1. Patients_Master**
```DAX
Total Patients = COUNTROWS(Patients_Master)
%Female = DIVIDE(CALCULATE([Total Patients], Patients_Master[Gender]="Female"), [Total Patients])
Avg Age = AVERAGE(Patients_Master[Age])
```
**2.  Encounters_Log & Charges**
```DAX
Total Encounters = DISTINCTCOUNT(Encounters_Log[EncounterID])
Total Revenue = SUM(Encounters_Log[ChargeAmount])
Avg Charge = DIVIDE([Total Revenue], [Total Encounters])
Charge Tier = SWITCH(TRUE(), Encounters_Log[ChargeAmount] < 430, "Low", Encounters_Log[ChargeAmount] < 470, "Mid", "High")
```
**3.  Insurance_Claims - Revenue Cycle**
```DAX
Total Billed = SUM(Insurance_Claims[AmountBilled])
Total Paid = SUM(Insurance_Claims[AmountPaid])
Total Remaining = SUM(Insurance_Claims[Remaining])
Collection Rate % = DIVIDE([Total Paid], [Total Billed], 0)
Payment% = DIVIDE(Insurance_Claims[AmountPaid], Insurance_Claims[AmountBilled], 0)
```
**4. Pharmacy & Prescriptions**
```DAX
Total RX Charges = SUM(Prescriptions[Charges])
Avg RX Charge = DIVIDE([Total RX Charges], [Total Prescriptions])
Medication Category = SWITCH(TRUE(), CONTAINSSTRING(Prescriptions[Medication Name], "Amoxicillin"), "Antibiotic",...)
```
**5. Vitals_Log - Clinical Risk**
```DAX
Avg Systolic = AVERAGE(Vitals_Log[SystolicBP])
BP Category = SWITCH(TRUE(), Vitals_Log[SystolicBP] < 120 && Vitals_Log[DiastolicBP] < 80, "Normal",..., "Hypertension Stage 2")
MAP = Vitals_Log[DiastolicBP] + (Vitals_Log[SystolicBP] - Vitals_Log[DiastolicBP]) / 3
% Stable = DIVIDE(CALCULATE(COUNTROWS(Vitals_Log), Vitals_Log[Risk_Status]="Stable"), [Total Vitals])
```
---
## 📊 Dashboard Visuals
The Power BI dashboard includes seven visuals covering:
- Providers & Facility performance  
- Charges  and Pharmacy  distribution  
- Procedure category analysis  
- Vitals and Risk Overview status  
- Pharmacy Performances by reason  
- Financial and Operational metrics effectiveness  
- Billing Metrics  analysis (forecasted Providers staffing  and procedural growth)
  
  ---

  ## 📈 Key Insights
- Dataset has 1,000 patients with avg age 50.3 years. Gender: Female 365, Male 319, Non-binary 316. 
- Encounter revenue: $1,507,649.02 from 5,491 encounters.
- Claims: $1,386,950.91 billed, $1,109,467.16 paid, $277,483.75 remaining (80% collection rate). 
- Pharmacy: 2,656 prescriptions worth $736,194.88.  
- Office Visits is the highest-revenue encounter category at ∼$394,938.27; Telehealth and Hospital Admission also contribute high.  
- CPT 80053 is the highest-revenue CPT code at ∼$391,616.27.
- Diagnosis E11.9 (Type 2 Diabetes) is the highest-revenue diagnosis at ∼$316,422.20.
- Dr. James Singh is the highest-revenue provider at ∼$111,400.46.
- Entire insurance summary is associated with Aetna only.

  ---

  ## 🔮 Predictions
- Historical monthly encounter, charge and prescription trends can be used as a baseline for forecasting demand and revenue for next quarter.
- Repeated vitals and diagnosis patterns (e.g., Hypertension Stage 1 = 169 patients) can flag patient cohorts that may require closer monitoring and preventive care.
- High-cost encounter/service categories (Hospital Admission) can be monitored to anticipate resource and revenue-cycle pressure.
- Seasonal spikes in Office Visits can be predicted for capacity planning.

  ---

  ## ✅ Recommendations
- Prioritize High-Revenue Services: Focus capacity and coding-quality review on Office Visits and CPT 80053.
- Improve Revenue Cycle: Investigate outstanding balance of $277K and payment leakage to improve collection performance from 80% to >95%.
- Workforce Planning: Use provider/specialty revenue and staffing indicators (High Staffing vs OK) to align workforce capacity with demand.
- Pharmacy Optimization: Use pharmacy-level (CVS, Walgreens, Hospital) and medication-level revenue patterns to optimise inventory and formulary planning.
- Population Health Program: Use demographic and clinical risk segmentation (Age Group: Senior Citizen is high risk, BP Category) to support targeted preventive-care and follow-up programmes.
  
---
## 📌 Conclusion
The integration of Excel and Power BI provides an effective end-to-end healthcare analytics workflow, from structured data preparation and PivotTable validation through DAX-based modelling and interactive dashboard reporting. The project demonstrates how patient, clinical, operational and financial data can be consolidated into a single decision-support framework.
The resulting Power BI report supports descriptive, diagnostic, predictive and prescriptive analysis, while the Excel workbook provides transparent source-level calculations and validation summaries. This solution empowers healthcare administrators to make data-backed decisions for better patient care and financial performance.

---

## 🙏 Acknowledgements
- Tools: Microsoft Excel, Power BI Desktop.
- Data Source: ExcelX.com for providing The Ultimate Healthcare Dataset Generator for Excel & Power BI.
- Guidance from Entri Elevate - Program In AI Driven Data Analysts.
---

## 📜 License
This project is for educational purposes only. The dataset is synthetic/simulated for learning Power BI and Excel analytics. You are free to use the DAX logic and dashboard structure with attribution.
