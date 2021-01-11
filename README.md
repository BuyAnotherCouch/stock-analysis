# Overview of Project
Steve is a finance professional who is looking to help his parents invest into the green energy sector.  While his parents are solely interested in DAQO stocks, Steve would also like to know how other stocks in the sector are performing. Using the stock data that he provided, we are tasked with analyzing the data to come up with the performance data needed to decide on which stocks to invest in. 

# Results
The raw data provided included performance data for twelve green energy stocks for the years 2017 and 2018.

To kick off our analysis, we created a macro in VBA, made up of a series of for loops, to return stock data for DAQO.

We then applied the same logic to the entire array of stocks to return stock data for all stocks.

Thanks to the macros that we built, we were able to determine that all stocks but TERP, generated a return in 2017, with DQ, SEDG, ENPH and FSLR leading in performance. 2018, however, presented an entirely different result with only ENPH and RUN generating returns.

# Summary
While the macros we built provided Steve with the ability to quickly generate the insights needed to decide on which stocks to invest for his parents, the processing time for our macros highlighted a limitation that we needed to address. If in the future, Steve wanted to add extra datasets to the workbook, the macro would take longer to run. To address this, we refactored our code to generate the same results with fewer loops. 

In this project, we were able to reduce processing times from around 0.97 seconds to 0.21 seconds for our 2017 stock data, and around 1.01 seconds to 0.20 seconds for our 2018 data. 

This not only improved our processing speeds, but it also made our code much cleaner and easier to read – which also makes it easier to edit in the future, if needed.

One disadvantage to refactoring is that it may take a lot of time to clean up large amounts of code, which may be an issue you are short for time or concerned about budgets.
