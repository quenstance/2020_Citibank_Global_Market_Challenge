# Portfolio Optimization Macro

### Overview

This **VBA (Visual Basic for Applications)** macro automates a complete **financial analysis workflow** within Excel. It's designed to streamline the process of downloading, organizing, and analyzing financial data to construct an optimized investment portfolio. The project was originally developed for the **Citibank Global Markets Challenge 2020** case competition. More details are available on my portfolio website: [https://quenstance.pages.dev/projects/2020-03-citibank-gmc](https://quenstance.pages.dev/projects/2020-03-citibank-gmc).

---

### Key Features

* **Automated Data Retrieval & Consolidation**: Downloads historical stock and index data from Yahoo Finance. It then automatically cleans and collates adjusted close prices into a single sheet, preparing the data for analysis.
* **Dynamic File Management**: Provides user-selectable options to either export processed data to CSV files or collate it within the macro workbook.
* **Mean-Variance Optimization**: Uses Excel's Solver to perform portfolio optimization. The code finds the minimum variance portfolio for a target return, a core concept of Modern Portfolio Theory. This process involves setting several constraints on asset weights, such as a **"no short selling"** restriction (weights must be non-negative) and limits on exposure to specific asset classes. The results of these optimizations are used to plot the **Efficient Frontier**, a graph that shows the optimal portfolios for a given level of risk. 

---

### How to Use

The macro workbook, "**2020 Citibank GMC Campus Final**," is a self-contained project that can be run directly within Excel. The accompanying `.cls` and `.bas` files are provided for code review and version control.

#### **Step 1: Enable the Solver Add-in**

This macro requires the **Solver Add-in**. If it is not already enabled, go to **File > Options > Add-ins**, select **Excel Add-ins** from the dropdown menu, click **Go**, and then check the box for **Solver Add-in**.

#### **Step 2: Set up Your Data**

Navigate to the **GetData** worksheet. Enter the stock ticker symbols and desired parameters (dates, frequency) directly into the designated cells.

#### **Step 3: Run the Macro**

The project is organized into separate modules, which you will need to run sequentially to complete the full workflow. For a detailed breakdown of each module and how to run them, refer to the [TECHNICAL_DETAILS.md](TECHNICAL_DETAILS.md) file.

1.  **Run `Module 1: DownloadData`** to retrieve and consolidate the data.

2.  **Run `Module 2: CleanData`** to prepare the consolidated data. This macro will automatically **remove all formulas**, replacing them with static values. It will also find and **delete any "null" values and empty rows** to ensure the dataset is dense and ready for calculation.

3.  **Run `Module 3: ExtractRf`** to isolate the risk-free rate data. This macro will create a new sheet and copy the risk-free rate data to it, formatting it as a percentage for use in subsequent portfolio calculations.

4.  **Run `Module 4: RestrictHoldings` or `RestrictShortSales`** to perform the optimization.

#### **Step 4: Review Your Results**

The optimization results, including the **efficient frontier**, will be displayed on the **MVF (Mean-Variance Frontier)** sheet.

---

### Project Status & Known Limitations

This repository contains the original files submitted for the **Citi Global Markets Challenge 2020**. The code is presented as a historical record of the competition submission and has not been updated. As such, it contains several limitations based on the technology and data available at the time.

* **Module 1 - Data Retrieval**: The data retrieval process has several limitations:
    * **Windows-Only Compatibility**: The macro uses `WinHttp.WinHttpRequest.5.1`, a **COM object** exclusive to the Windows operating system, making it incompatible with macOS.
    * **Fragile Web Scraping**: The code relies on searching for specific string patterns to extract data (e.g., `"""crumb""":""`). This is a very **fragile method** for web scraping, as websites frequently update their code, which may cause the macro to fail. A more robust solution would be to use a dedicated data retrieval library like Python's `yfinance` or a consistent financial data API.
    * **File Deletion Logic**: The `Module 1` is designed to populate and then remove old data by deleting all sheets that occur after the **"GetData"** sheet. This is a crucial part of the workflow and ensures the new data is added to a clean workbook for the portfolio analysis.
* **Risk-Free Rate Data**: The model uses the `Adjusted Close Price` data field from Yahoo Finance for the `^IRX` (13-week Treasury Bill) as a proxy for the risk-free rate. This is a known technical inaccuracy, as this field is a remnant of the equity data structure. A more reliable and technically sound approach for a portfolio optimization model would be to use **actual yield data** directly from sources like the U.S. Department of the Treasury or a financial data API that provides bond yields.

---

### Author

**Quenstance Lau**
* **Web Portfolio**: https://quenstance.pages.dev/
* **LinkedIn**: https://www.linkedin.com/in/quenstance/
* **GitHub**: https://github.com/quenstance
