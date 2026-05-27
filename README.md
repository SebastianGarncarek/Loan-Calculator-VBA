# Loan Calculator

## Project Overview
This project presents an advanced financial simulation engine and interactive dashboard built in Excel VBA designed to model and calculate credit repayment schedules. 

Rather than just calculating standard static loans, a significant portion of this project focuses on handling dynamic risk by processing historical market data. The application allows users to seamlessly simulate and compare both fixed and floating interest rate scenarios using live-updating calculation pipelines.

<img width="423" height="740" alt="image" src="https://github.com/user-attachments/assets/dd9e0ed5-e744-4c47-a4c7-b85e466b4c05" />

## Key Highlights
The core strength of this application lies in its production-grade interest rate calculation engine. When dealing with floating-rate loans, the algorithm dynamically computes a **60-day moving average of the WIBOR benchmark** based on a historical database ranging from 2020 to 2024.

Instead of crashing or freezing when encountering non-trading days, weekends, or data gaps, the backend implements a robust, defensive search mechanism. It automatically backtracks to the closest available historical trading record and leverages explicit data-type validation to filter out hidden text cells or formula errors before calculating financial metrics.

To ensure data integrity, the application enforces strict structural protection. The reporting worksheet automatically locks into a read-only state upon schedule generation, preventing users from accidentally modifying numbers or ruining the table architecture, while still allowing manual cell-width adjustments.

<img width="854" height="598" alt="image" src="https://github.com/user-attachments/assets/da58dacf-ddba-4b79-9628-11c32864d8da" />

## Methodology
The application architecture is cleanly decoupled into UI events and calculation logic, following a standard financial engineering approach:
1. **Database & Range Querying:** Dynamically locating the dataset boundaries using safe `lastRow` detection methods rather than hardcoded ranges.
2. **Data Cleansing & Filtering:** Iterating through the rolling average window with explicit error trapping (`IsError`) and numeric validation (`IsNumeric`).
3. **Amortization Scheduling:** Running multi-variable loops to dynamically decrease principal balances, adjust the final installment to clear residual fractions, and print formatted outputs.

## Code & System Architecture
The repository is structured to allow direct code review within GitHub without needing to download the workbook:
* `CalculatorUI.frm` / `.frx` – The graphical user interface handling real-time asynchronous input validation and field constraints.
* `LoanLogic.bas` – The isolated core financial module compiled under strict variable declaration modes (`Option Explicit`).
* `LoanInstallmentsScheduler.xlsm` – The macro-enabled Excel workbook containing the dashboard, user triggers, and the historical WIBOR database.

## 📖 How to Run
To test the interactive elements of the calculator locally on your machine:
1. Download the `LoanInstallmentsScheduler.xlsm` workbook from this repository.
2. Open the file in Microsoft Excel and click **Enable Macros** / **Enable Content** on the top yellow warning bar.
3. On the `Main` worksheet, click the primary launch button to initialize the custom graphical interface.
4. Input your custom loan terms to dynamically overwrite and generate a fresh amortization report.

## Technologies Used
* **Language:** VBA (Visual Basic for Applications)
* **Environment:** Microsoft Excel
* **Core Concepts:** Object-Oriented UserForms, Dynamic Range Manipulation, Financial Engineering Algorithms, Defensive Programming, Data Validation.

## Author
**Sebastian Garncarek** * M.Sc. Student in Computational Mathematics
* [Connect with me on LinkedIn](https://www.linkedin.com/in/sebastian-garncarek-a0376a405)
