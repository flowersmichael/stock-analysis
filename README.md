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

As you can see, the refactored code was more than five times faster than our original code!

## Summary
Refactoring, or editing, is an important part of the coding process. When we refactor code, the output looks the same. We preserve the functionality while improving the code design, structure, or implementation. Some reasons to refactor code include greater effiency or enhanced readability.

Sometimes it might not be necessary to refactor code, for example if we are going to be the only user and it's a one-off task, or if the code is really short. Oftentimes, however, it can be critical to refactor code. If we are performing a repetitive task thousnds of times or more, making a code even slightly more efficient can make a significant difference. It's also important to refactor a code for readability purposes, in order to make it easier to undersand for other users or even our future selves.

In this specific case, our original code worked well when we were only running it on 12 stocks. The speed difference wasn't noticeable. If we want to reuse our code to analyze hundreds or thousands of stocks, however, the efficiency will matter and the speed difference will be apparent. By using multiple arrays, we were able to loop through the code only one time, thus making the process more than 5X faster.





