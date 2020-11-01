## Overview of Project

The project is to automate tasks for calculations and formatting of a stocks volume and price dataset in Excel, using VBA Macros.

### Purpose

The purpose of the project is twofold -
* to conduct analysis of 12 different stocks for two different years in order to inform investment decisions for prospective investors and then,

* refactor the code for reusability and extend functionality so the macro can perform stocks analysis on variable number of stocks, and calculate stock volumes and return rate; and also increase performance by improving run time.

## Results

### Stock analysis for 2017

* In 2017 the stock with almost 200% return was DQ, followed closely by a high return of 184.5% of SEDG. The stocks ENPH and FSLR performed very well too with a return exceeding 100%. The stocks SPWR and FSLR were the top two in terms of total daily volume.

![Stocks analysis 2017](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/stocks-analysis-2017.png)


### Stock analysis for 2018

* In 2018 the stocks with over 80% return were ENPH and RUN. All other stocks dropped to a negative return percentage with the lowest being - DQ at -62.6% and JKS at -60.5%. The stocks ENPH, RUN and SPWR were the top three in terms of total daily volume.

![Stocks analysis 2018](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/stocks-analysis-2018.png)



### Code execution time

* The code execution time decreased significantly for both stocks analysis of 2017 and 2018.

* Run time (0.64s) before and after (0.57s) refactoring for 2017.

![runtime before - 2017](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/runtime-before-refactor-2017.png)


![runtime after - 2017](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/runtime-after-refactor-2017.png)



* Run time (1.07s) before and after (0.67s) refactoring for 2018.

![runtime before - 2018](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/runtime-before-refactor-2018.png)


![runtime after - 2018](https://github.com/divitaN-dev/stocks-analysis-macro/blob/main/resources/runtime-after-refactor-2018.png)



## Summary

### Refactoring Code and implications on the current project

* Refactoring code can help remove code redundancy, remove duplicate variables and hard coded variable values and make them dynamic, and also assign proper data types if that was previously not the case.

* For this project the variables were reassigned as arrays and the calculation for stock volumes and return percentages were stored in array indices. This enabled the results of all analysis to be written using a single loop and consequentially descreased code run time significantly.

* However, refactoring code does come with its unique challenges. It can intially be difficult to understand someone else's logic behind the sequence of operations, this can be alleviated to some extend by using comments. It also requires a bit of trial and error to make sure a refactored statement doesn't inadvertently change or break the intented functionality of the code.

* For this project it took some trial and error to ensure the array indices corresponded to the correct stock volumem and return percentages on the output excel sheet.



