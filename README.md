# Fraud Detection Analysis Using Microsoft Excel

<img width="1208" height="397" alt="image" src="https://github.com/user-attachments/assets/f3b1066f-1888-4743-a745-333ac0308a58" />
<img width="1127" height="300" alt="image" src="https://github.com/user-attachments/assets/ad39f135-7f3c-4445-82d4-22c7e71ff731" />
<img width="887" height="642" alt="image" src="https://github.com/user-attachments/assets/71866dce-86f6-473e-9442-637dd408a0b4" />



End-to-End Data Analytics Project Documentation

Project Type: Financial Transaction Fraud Detection\
Tool: Microsoft Excel\
Dataset: 1,500 synthetic transactions\
Analysis Period: January–October 2026\
Methodology: Data Cleaning, Exploratory Data Analysis (EDA), Rule-Based Fraud Detection, Risk Scoring, KPI Analysis, and Dashboard Reporting.

## 1. Project Overview

### 1.1 Introduction

Fraud Detection Analysis is a financial data analytics project designed to identify potentially suspicious transactions, analyze fraud patterns, and generate actionable business insights to help reduce financial losses.

This project uses Microsoft Excel to analyze transaction data and classify transactions based on predefined risk indicators.

The dataset contains 1,500 synthetic financial transactions, including both fraudulent and legitimate transactions.

### 1.2 Business Problem Statement

Financial institutions and payment companies process thousands of transactions every day. Some transactions may be fraudulent, resulting in financial losses and customer dissatisfaction.

Manually reviewing every transaction is inefficient and time-consuming.

Business Question:

How can a financial institution identify potentially fraudulent transactions using transaction behavior, customer information, and risk indicators while minimizing false positives?

### 1.3 Project Objectives

1. Identify fraudulent and legitimate transaction patterns.
2. Analyze transaction amounts, payment methods, merchant categories, and geographic indicators.
3. Develop Excel-based fraud detection rules.
4. Calculate a risk score for each transaction.
5. Classify transactions into Low, Medium, and High risk categories.
6. Evaluate fraud detection performance using analytical metrics.
7. Develop an executive dashboard for fraud monitoring.
8. Recommend strategies to improve fraud investigation and monitoring.

## 2. Dataset Understanding

The project uses a synthetic dataset containing 1,500 transactions.

Total Transactions

# 1,500

Fraud Cases

# 122

Legitimate Cases

# 1,378

Fraud Rate

# 8.13%

These figures describe the generated practice dataset, not real-world financial transactions.

### 2.1 Dataset Columns

The Transactions sheet contains 26 columns divided into three categories.

A. Transaction and customer information

| Column                | Description                                      |
| --------------------- | ------------------------------------------------ |
| Transaction_ID        | Unique transaction identifier                    |
| Transaction_Date      | Date and time of the transaction                 |
| Customer_ID           | Unique customer identifier                       |
| Merchant_ID           | Merchant identifier                              |
| Amount_INR            | Transaction amount in Indian rupees              |
| Merchant_Category     | Category such as Electronics, Fashion, or Travel |
| Country               | Country associated with the transaction          |
| Payment_Method        | Credit Card, Debit Card, UPI, etc.               |
| Device_Type           | Mobile, Desktop, or Tablet                       |
| IP_Risk_Score         | Synthetic IP risk indicator between 0 and 1      |
| Distance_From_Last_KM | Distance from the previous transaction location  |
| Account_Age_Days      | Age of the customer account in days              |
| Failed_Attempts       | Number of recent failed payment attempts         |
| Transaction_Hour      | Hour of the transaction, from 0 to 23            |
| Actual_Label          | Known synthetic label: Fraud or Legitimate       |

B. Fraud detection features

| Column                 | Description                                                     |
| ---------------------- | --------------------------------------------------------------- |
| High_Amount_Flag       | Identifies transactions of ₹30,000 or more                      |
| High_IP_Risk_Flag      | Identifies IP risk scores of 0.70 or more                       |
| Long_Distance_Flag     | Identifies distances of 150 km or more                          |
| New_Account_Flag       | Identifies accounts aged 30 days or less                        |
| Failed_Attempts_Flag   | Identifies three or more failed attempts                        |
| Night_Transaction_Flag | Identifies transactions at or before 4 AM, or at or after 11 PM |
| Rule_Flag_Count        | Counts the triggered risk indicators                            |

C. Risk assessment outputs

| Column            | Description                                                                         |
| ----------------- | ----------------------------------------------------------------------------------- |
| Risk_Score        | Weighted risk score from 0 to 100                                                   |
| Risk_Band         | Low, Medium, or High                                                                |
| Model_Decision    | Approve or Review recommendation                                                    |
| Correct_Detection | Evaluation outcome: True Positive, False Positive, False Negative, or True Negative |

## 3. Data Cleaning and Preparation

Objective: Ensure data quality, consistency, and accuracy before performing analysis.

### Step 1: Check for duplicate transactions

```
=COUNTIF($A$2:$A$1501,A2)
```

If the result is greater than 1, the Transaction_ID appears more than once.

### Step 2: Identify missing values

```
=COUNTBLANK(A2:O1501)
```

Check important fields individually, particularly Transaction_ID, Transaction_Date, Amount_INR, and Actual_Label.

### Step 3: Validate transaction amounts

```
=IF(E2<=0,"Invalid","Valid")
```

### Step 4: Validate IP risk scores

```
=IF(AND(J2>=0,J2<=1),"Valid","Invalid")
```

### Step 5: Validate transaction dates

Ensure Transaction_Date is stored as a valid Excel date and time value rather than text.

### Step 6: Check data consistency

- Transaction IDs should be unique.
- Transaction amounts should be positive.
- IP risk scores should be between 0 and 1.
- Failed attempts and account ages should be nonnegative.
- Transaction hours should be between 0 and 23.
- Fraud labels should contain only `Fraud` or `Legitimate`.

Expected Output: A validated dataset ready for analysis.

## 4. Exploratory Data Analysis (EDA)

EDA helps identify suspicious patterns, anomalies, and potential fraud indicators.

### 4.1 Transaction Volume Analysis

Business Question: How many transactions were processed?

```
=COUNTA(Transactions!A2:A1501)
```

Expected result: 1,500 transactions.

### 4.2 Fraud Transaction Analysis

Business Question: How many transactions are fraudulent?

```
=COUNTIF(Transactions!O2:O1501,"Fraud")
```

Expected result: 122 fraud transactions.

### 4.3 Fraud Rate Analysis

```
=COUNTIF(O2:O1501,"Fraud")/COUNTA(A2:A1501)
```

Expected result: 8.13%.

Business Interpretation: Approximately 8 out of every 100 transactions in this synthetic dataset are labeled fraudulent.

### 4.4 Merchant Category Analysis

Business Question: Which merchant categories have the highest fraud rate?

Create a PivotTable using:

| PivotTable Area | Field                   |
| --------------- | ----------------------- |
| Rows            | Merchant_Category       |
| Columns         | Actual_Label            |
| Values          | Count of Transaction_ID |

Calculate the fraud rate:

\\[ \text{Fraud Rate}=\frac{\text{Fraud Transactions}}{\text{Total Transactions}}\times100 \\]

Recommended visualization: a column chart for fraud counts and a separate chart for fraud rates.

### 4.5 Payment Method Analysis

Business Question: Which payment methods have higher fraud rates?

Analyze Credit Card, Debit Card, UPI, Net Banking, and Digital Wallet transactions.

Compare fraud rates rather than just fraud counts, because each payment method may have a different transaction volume.

### 4.6 Time-Based Fraud Analysis

Business Question: During which hours do fraudulent transactions occur most frequently?

Create a PivotTable using:

- Rows: Transaction_Hour
- Columns: Actual_Label
- Values: Count of Transaction_ID

Recommended visualization: a column chart or line chart.

### 4.7 High-Value Transaction Analysis

Analyze transaction amounts to determine whether high-value transactions have greater fraud exposure.

Compare average transaction amount, median amount, maximum amount, and fraud rate across transaction amount bands.

## 5. Fraud Detection Rule Development

The project uses six risk indicators to calculate a weighted score.

| Risk Indicator          | Condition             | Points |
| ----------------------- | --------------------- | ------ |
| High transaction amount | Amount ≥ ₹30,000      | 15     |
| High IP risk            | IP risk score ≥ 0.70  | 25     |
| Long distance           | Distance ≥ 150 km     | 20     |
| New account             | Account age ≤ 30 days | 15     |
| Failed attempts         | Failed attempts ≥ 3   | 15     |
| Night transaction       | Hour ≤ 4 or ≥ 23      | 10     |
| Maximum score           |                       | 100    |

### 5.1 Risk Score Formula

```
=MIN(100,ROUND(
P2*15+
Q2*25+
R2*20+
S2*15+
T2*15+
U2*10,0))
```

Here, columns P to U contain the six binary risk flags.

### 5.2 Risk Classification

| Risk Score | Risk Band |
| ---------- | --------- |
| 0–34       | Low       |
| 35–59      | Medium    |
| 60–100     | High      |

Excel formula:

```
=IF(W2>=60,"High",
 IF(W2>=35,"Medium","Low"))
```

### 5.3 Decision Logic

```
=IF(W2>=60,"Review","Approve")
```

A Review decision means the transaction requires further investigation. It does not prove that fraud has occurred.

## 6. Fraud Detection Performance Evaluation

A confusion matrix is used to evaluate the performance of the rule-based detection system.

|                   | Actual Fraud        | Actual Legitimate   |
| ----------------- | ------------------- | ------------------- |
| Predicted Review  | True Positive (TP)  | False Positive (FP) |
| Predicted Approve | False Negative (FN) | True Negative (TN)  |

### 6.1 Precision

\\[ \text{Precision}=\frac{TP}{TP+FP}\times100 \\]

Measures the percentage of reviewed transactions that are actually fraudulent.

### 6.2 Recall

\\[ \text{Recall}=\frac{TP}{TP+FN}\times100 \\]

Measures the percentage of all actual fraud transactions that were flagged for review.

### 6.3 Accuracy

\\[ \text{Accuracy}=\frac{TP+TN}{TP+FP+FN+TN}\times100 \\]

Measures the percentage of all transactions classified correctly.

### 6.4 False Positive Rate

\\[ \text{FPR}=\frac{FP}{FP+TN}\times100 \\]

Measures the percentage of legitimate transactions incorrectly flagged for review.

Business Importance: High recall helps reduce missed fraud, while high precision helps minimize unnecessary investigations. A practical fraud detection system must balance both according to business costs.

## 7. Executive Dashboard Development

The workbook includes an Executive Dashboard sheet with KPI cards and an Analysis sheet with category-level charts. Additional PivotCharts and slicers can further improve interactivity.

### 7.1 Recommended Dashboard KPIs

- Total Transactions
- Fraud Cases
- Fraud Rate
- Total Transaction Value
- Transactions Flagged for Review
- Review Precision
- Fraud Recall
- False Negatives

### 7.2 Recommended Visualizations

Executive KPI Overview

Display transaction volume, fraud rate, flagged transactions, and detection performance.

[Fraud Detection in Financial Transactions: Insights, Challenges, and Strategic Recommendations | by Helen-Nellie Adigwe | Medium](https://images.openai.com/static-rsc-4/46MdtQZBZgM0ngwf8u8e0v8BE89ccn-AnqZJddv8CBdBT8ZwO32wtdKGB9dmrqbDvGJkzIdFhB9DUGluDfCroNObJ4sdeG83L80f60CmKjHplkznOuSu1qRyWW6azjdgti930YPcAOzOYyx8mkyZM7tcX_7a35tbxRK1EH4YD3U?purpose=inline)

Fraud by Merchant Category

Compare fraud counts and fraud rates across different merchant categories.

[インターネットで不正行為を行うユーザーの特徴をグラフ化するとこうなる - GIGAZINE](https://images.openai.com/static-rsc-4/bPyU3l7WzGj9IlGD-KwfUXWYtNjeJ381lHjALHxi1xzKzE6gQl3uQmxoqB6SwiJgZsQom6egC1YgRdRUJUE3OKn__Y3XOedwLfhSJJc9cASIPeiSusvC7Ym0C4gdIHnBWb42YmjDb66k7Ai7htt4SaVtqzSd-ilmkguFbyGZjMg?purpose=inline)

Hourly Fraud Trends

Examine transaction patterns throughout the day and identify potentially unusual periods.

[Projects | Michael Duggan](https://images.openai.com/static-rsc-4/1ewJ_pX8q-K0uJ11cdihEwv_8O-CQlP7XlWfHklnosLZ2jMsb9kAVKuVlW2p398jCcrqwaqHu78jGn5INvqSlus3DjWiV4u_g018HIn7KWJmGovUKJ-6bgBSEv444On-5zIq2sDS_5VMG4WUV46sykqlgYjiVMi9zBCYw5timBc?purpose=inline)

Risk Distribution

Compare Low, Medium, and High risk transactions and prioritize investigation resources.

### 7.3 Recommended Dashboard Filters

Add Excel slicers for Merchant_Category, Payment_Method, Country, Risk_Band, and Device_Type. Add a timeline filter for Transaction_Date where supported.

Note: The images above are visual references, not screenshots of the actual workbook.

## 8. Business Insights

The following findings are verified from the generated dataset or workbook structure.

Insight 1 — Fraud prevalence: Out of 1,500 transactions, 122 are labeled fraudulent, resulting in a fraud rate of 8.13%.

Insight 2 — Class imbalance: Legitimate transactions represent 91.87% of the dataset. Accuracy alone is therefore insufficient to evaluate fraud detection performance.

Insight 3 — Investigation workload: The generated investigation queue contains 89 transactions flagged by the initial scoring threshold, approximately 5.93% of all transactions.

Insight 4 — Multiple risk signals: The project combines transaction amount, IP risk, location distance, account age, failed attempts, and transaction timing.

Further analysis is required before stating which merchant category or payment method has the highest fraud rate. Precision, recall, and false-positive results should also be taken from the recalculated workbook.

## 9. Business Recommendations

1. Prioritize high-risk transactions: Investigate transactions scoring 60 or more, particularly those with high potential financial exposure.
2. Monitor unusual account activity: Combine signals such as new accounts, repeated failed attempts, and suspicious IP behavior.
3. Reduce false positives: Evaluate alternative risk thresholds before implementing stricter controls.
4. Monitor category-specific risk: Use fraud rates and potential financial exposure to identify categories requiring closer investigation.
5. Track detection performance: Monitor Precision, Recall, False Positive Rate, and the value of missed fraud over time.
6. Maintain human verification: Do not automatically reject transactions solely because of these demonstration rules.

## 10. Project Limitations

- The dataset is synthetic and does not represent real financial transactions.
- The fraud labels and risk patterns were generated for practice.
- The detection system uses predefined rules rather than a trained machine-learning model.
- IP risk scores and distance measurements are synthetic indicators, not independently verified security signals.
- Real-world deployment would require historical behavioral data, secure data handling, validation against investigation outcomes, and continuous monitoring.

## 11. Project Conclusion

The Fraud Detection Analysis project demonstrates how Microsoft Excel can transform transaction data into risk-based business insights.

It covers data understanding, validation, exploratory analysis, feature engineering, risk scoring, classification evaluation, dashboard reporting, and investigation prioritization.

Business Value: The project provides a transparent framework for identifying potentially suspicious transactions and supporting fraud investigation decisions.

Technical Skills Demonstrated: Microsoft Excel, Data Cleaning, PivotTables, IF, COUNTIF, COUNTIFS, SUMIFS, Conditional Formatting, KPI Development, Risk Scoring, Confusion Matrix, and Dashboard Design.

## 12. Interview Explanation

**Project: Fraud Detection Analysis Using Microsoft Excel**

I developed an end-to-end Fraud Detection Analysis project using Microsoft Excel.

The project involved analyzing 1,500 synthetic financial transactions to identify suspicious transaction patterns and evaluate fraud detection rules.

I worked on data understanding, data validation, exploratory data analysis, and feature engineering.

I developed a rule-based fraud risk scoring system using transaction amounts, IP risk scores, location distance, account age, failed attempts, and transaction timing.

Based on the calculated risk scores, I classified transactions into Low, Medium, and High risk categories and identified transactions requiring further investigation.

I also developed an executive dashboard to monitor transaction volumes, fraud rates, flagged transactions, and detection performance using Precision, Recall, and Accuracy.

This project strengthened my Excel, analytical thinking, KPI development, and business problem-solving skills.

You can use this documentation as your project report, portfolio reference, and interview preparation material.
