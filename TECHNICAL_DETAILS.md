# Technical Details

This document provides a detailed breakdown of the VBA macro workbook's architecture and functionality.

---

### Project Modules

#### Sheet 3: "GetData"

This code is an **event procedure** that triggers automatically when a change is made on the **"GetData" worksheet**. The `Worksheet_Change(ByVal Target as Range)` sub-routine uses the `Target` object to identify the modified cell. The code's primary function is to apply and clear formatting, using a `With...End With` block to manipulate the `Interior` and `Borders` properties of a specified `Range` object. It sets properties like `Pattern`, `TintAndShade`, `LineStyle`, and `Weight` to visually delineate the input area for stock tickers.

#### Module 1: Download Data

This is a **standard module** containing a collection of sub-routines and functions. The main routine, `DownloadData()`, orchestrates the entire process:

* **Application Control**: It disables `Application.Calculation`, `Application.ScreenUpdating`, and `Application.DisplayAlerts` to enhance performance during the macro's execution.
* **Data Retrieval**: It reads user-defined parameters such as `startDate`, `endDate`, and `frequency` from the "GetData" sheet. It converts calendar dates into **Unix time** (the number of seconds since January 1, 1970) for API requests.
* **Web Scraping**: The code uses a `WinHttp.WinHttpRequest` object to send `GET` requests to Yahoo Finance's API. It first calls `getCookieCrumb()` to obtain a session `cookie` and a `crumb`, which are required for authentication to download data.
* **Data Processing**: The `getYahooFinanceData()` sub-routine downloads the data as a **CSV (Comma-Separated Values)** string. It then uses the `Split()` function to parse this string into a **dynamic array** (`resultArray`). The `UBound()` and `ReDim Preserve` statements handle arrays of varying sizes. This array is then written directly to a new worksheet using `Resize()` to fit the data dimensions.
* **Data Sorting**: The `SortByDate()` sub-routine uses the `.SortFields.Add` method to sort the data on the "Date" column (`Range("A" & firstRow & ":A" & lastRow)`), either in `xlAscending` or `xlDescending` order based on user selection.
* **Output Routines**: The module includes `CopyToCSV()` to export individual worksheets as CSV files and `CollateData()` to combine the `Adjusted Close Price` data from all stock sheets into a single, comprehensive worksheet. This is achieved using the `VLOOKUP` function via the `Formula` property.

#### Module 2: Clean Data

The `CleanData()` macro takes the consolidated data and refines it. It performs a **PasteSpecial** operation with `xlPasteValues` to remove all underlying formulas. It then uses the `Range.Replace` method to find and clear "null" values. Finally, it uses `SpecialCells(xlCellTypeBlanks)` to identify and delete any empty rows, ensuring a dense dataset for analysis.

#### Module 3: Extract Rf

This macro is a specialized cleaning routine for the risk-free rate data. It copies a specific range of data and pastes it into a new sheet named "Rf." It applies the same cleaning and blank-row deletion methods as `CleanData()`. It then calculates the risk-free rate as a percentage using a formula written to the cells and formatted with `Selection.NumberFormat = "0.00%"`.

#### Module 4: Solver

This module is dedicated to **portfolio optimization** using the Excel Solver add-in. This is a crucial part of **Modern Portfolio Theory (MPT)**.

* **Formulas**: The `RestrictHoldingFormulae()` sub-routine writes **array formulas** (`MMULT` for matrix multiplication) to the worksheet to dynamically calculate the portfolio's variance and expected return.
* **Solver Setup**: The `RestrictHoldings()` sub-routine uses a loop to find the optimal asset weights for a range of expected returns. In each iteration, it:
    * **Resets** Solver using `SolverReset`.
    * **Adds Constraints** using `SolverAdd`. Key constraints include setting the sum of weights equal to 1 (`Relation:=2, FormulaText:="1"`) and restricting asset-class weights within a defined range.
    * **Sets the Objective**: The `SolverOk` method defines the **optimization problem**. It sets the target cell to the portfolio's variance (`SetCell:=Range("$AL$13").Offset(i, 0)`), specifies the objective is to **minimize** this value (`MaxMinVal:=2`), and identifies the `ByChange` cells (the asset weights) that Solver can adjust. The `GRG Nonlinear` engine is chosen for this type of non-linear optimization.
* **Results Handling**: The `SolverSolve(True)` command executes the Solver. An `If...ElseIf...Else` block then checks the results of the Solver. If the solution is valid based on a set of criteria, it keeps the results with `SolverFinish KeepFinal:=1`. Otherwise, it discards them with `SolverFinish KeepFinal:=2`. The code also uses `SolverOptions AssumeNonNeg:=False` to allow for **short selling**, where asset weights can be negative.

---

### Project Status & Known Limitations

This documentation reflects the original files submitted for the **Citi Global Markets Challenge 2020**. The code is presented as a historical record of the competition submission and has not been updated.

* **Module 1 - Data Retrieval**: The data retrieval process has several limitations:
    * **Windows-Only Compatibility**: The macro uses `WinHttp.WinHttpRequest.5.1`, a **COM object** exclusive to the Windows operating system, making it incompatible with macOS.
    * **Fragile Web Scraping**: The code relies on searching for specific string patterns to extract data (e.g., `"""crumb""":""`). This is a very **fragile method** for web scraping, as websites frequently update their code, which may cause the macro to fail. A more robust solution would be to use a dedicated data retrieval library like Python's `yfinance` or a consistent financial data API.
    * **File Deletion Logic**: The `Module 1` is designed to populate and then remove old data by deleting all sheets that occur after the **"GetData"** sheet. This is a crucial part of the workflow and ensures the new data is added to a clean workbook for the portfolio analysis.
* **Risk-Free Rate Data**: The model uses the `Adjusted Close Price` data field from Yahoo Finance for the `^IRX` (13-week Treasury Bill) as a proxy for the risk-free rate. This is a known technical inaccuracy, as this field is a remnant of the data provider's equity data structure. A more reliable and technically sound approach for a portfolio optimization model would be to use **actual yield data** directly from sources like the U.S. Department of the Treasury or a financial data API that provides bond yields.