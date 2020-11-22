# VBA of Wall Street

## Overview of Project
In this project, we wanted to refactor some VBA code we'd previously written to make it run faster. The speed was not an issue for a small number of stocks, but could make a significant difference when analyzing hundreds or thousands of stocks over multiple years.

### Purpose and Background
Our friend wanted analyze green energy stocks. We had stock price and volume data in Excel for the years 2017 and 2018, and decided to use VBA to analyze a group of 12 green energy stocks. 

In our initial analysis, we were able to create a macro that pulled daily price and volume data from one worksheet, and populated another worksheet with annual returns as well as total volume for the year.

Our VBA code performed well, but we wanted to rewrite the code to make it run more efficiently, and thus allow us to expand our analysis to hundreds or thousands of stocks. In order to make our code run faster, we used multiple arrays.

## Results
Thankfully, we were successful in our effort to make our code more efficient.

### Run Time: Original Code
Here are the run times for the data from the years 2017 and 2018, using our *original* code:

![Original Code Run Time (2017)](https://github.com/flowersmichael/stock-analysis/blob/main/Resources/2017_original_run_time.png)

![Original Code Run Time (2018)](https://github.com/flowersmichael/stock-analysis/blob/main/Resources/2018_original_run_time.png)

### Run Time: Refactored Code
In contrast, here are the run times for 2017 and 2018 using our *refactored* code:

![Refactored Code Run Time (2017)](https://github.com/flowersmichael/stock-analysis/blob/main/Resources/2017_refactored_run_time.png)

![Refactored Code Run Time (2018)](https://github.com/flowersmichael/stock-analysis/blob/main/Resources/2018_refactored_run_time.png)



