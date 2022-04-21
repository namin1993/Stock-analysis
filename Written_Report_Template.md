# VBA Challenge

## Overview of Project

### Purpose
'''
This project is about how to program with VBA script through a scenario where we would like to assess which stocks on a portfolio were traded the most in a given year and which stocks yielded positive price returns from the begining to the end of a year.
'''

## Results

![2017 Stock Analysis Runtime](https://myoctocat.com/assets/images/base-octocat.svg)

![2018 Stock Analysis Runtime](https://myoctocat.com/assets/images/base-octocat.svg)

'''
For each year, an analysis was performed on all the stocks listed. I determined:
  1.) The total volume that all stocks were traded in a given year
  2.) The percent closing price return of a given stock from the beginning to the end of a year.

In 2017, SPWR was the highest traded stock with a total volume of 782,187,000. All stocks in the portfolio except for TERP yielded positive returns. Some stocks yielded extremely high returns such as DQ, EPNH, DSLR, and SEDG. 

In 2018, ENPH was traded the most with its volume of 607,473,500. ENPH and RUN were the only 2 stocks which yielded positive returns in 2018. JKS yielded the worst return in stock with a -60.5% return.
'''


## Summary

- What are the advantages or disadvantages of refactoring code?
  - Refactoring code can sometimes be a time-consuming task if you have to keep track of many variables, conditions, loops, and  general boolean or arithmetic logic within the script. However, good programming takes into consideration good performance and speed of a program to complete its execution since blocks of memory have to be set aside by the OS in order for the CPU to process all steps within a given program while also running several other system programs and applications in the background.

- How do these pros and cons apply to refactoring the original VBA script?
  - When refactoring my original VBA script, I had to combine all the inner code within the Cell Formatting subroutine and Run_Stock_Analysis button subroutine into the main Stock_Analysis subroutine. In my opinion, it was just another way of organizing code to allow for a faster runtime so that the VBA compiler did not need to go through too many steps in order to calculate the percent return and volume for a given stock with a given year.