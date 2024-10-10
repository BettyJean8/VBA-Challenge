# VBA-Challenge
Project Description
  This project focuses on using VBA scripting to analyze stock market data from multiple quarters. The script loops through stock data, calculates quarterly performance metrics for each stock, and outputs the     results to the worksheet. The goal is to automate the tedious process of stock market data analysis using VBA, including calculating percentage changes, stock volumes, and identifying the stocks with the greatest increases and decreases. The project is designed to demonstrate proficiency in VBA scripting and Excel automation.

Repository Structure
The repository contains the following files:

Britta NewlyMultiple_year_stock_data: This is the workbook for the assignment. 

Stock_Analysis_For: The main VBA script that performs the stock data analysis, loops through the quarters, and applies conditional formatting.

Great_Incrs_Great_Dcrs_Total_Vol: VBA script that performs finding the greatest increase in quarterly, the greatest decrease in quarterly change and the highest total stock volume.

Screenshots: A folder containing screenshots of the output results, demonstrating the successful execution of the script and the expected outcomes.

README.md: This README file that describes the project, its requirements, and instructions for use

Instructions
Script Features
Stock Data Analysis:

The script loops through all stock data for each quarter, gathering information such as ticker symbols, stock volume, opening prices, and closing prices.
It calculates the quarterly change and the percentage change from the opening price to the closing price for each stock.
The total volume for each stock is also calculated.
Outputs:

Ticker Symbol
Quarterly Change: The difference between the opening price at the start of the quarter and the closing price at the end of the quarter.
Percentage Change: The percentage change between the opening price and closing price.
Total Stock Volume: The total number of shares traded during the quarter.
Greatest Stock Performance:

The script identifies and displays the stock with the greatest percentage increase, the greatest percentage decrease, and the greatest total stock volume.
Conditional Formatting:

Positive quarterly changes are highlighted in green.
Negative quarterly changes are highlighted in red.

Multiple Worksheets:

The script runs across all worksheets in the workbook complining all data on the first worksheet of the workbook, ensuring it processes data from all quarters without manual intervention.
