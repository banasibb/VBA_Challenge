# Stocks Analysis VBA Challenge

## Overview of the Project
### Purpose
The purpose of the analysis requested by the client was to assess the performance of twelve different stocks to determine which performed best in 2017 and 2018. The purpose of the Module 2 challenge was to refractor this original code to improve its efficiency and reduce the amount of time needed to execute the code. 
### Data
The data set used in the analysis is organized by year into two separate tabs, "2017" and "2018," based on the stock issue date. The data includes performance of twelve different stocks by opening price, high and low values, closing price, adjusted closing price, and the volume of shares sold. Pertinent to this analysis are the open and close prices for each of the twelve different stocks, as well as the volume of shares sold. <br /><br />
To make the file easier to interact with for the end user, buttons were created to clear the analysis and to initialize the subroutine in which the user could enter the year they wish to perform the analyis on. Conditional formatting was used within the macro itself to highlight the top performing stocks. 

## Results
### Analysis
The screenshots of the refractoring exercise are below. 
Prior to refractoring the code, the run time for the 2017 analysis was 1.179 seconds, and the run time for the 2018 code was 1.121 seconds. These results can be found in the original [greenstocks.xlsm](https://github.com/banasibb/VBA_Challenge/blob/2dc24620f7c8c932b96ab481c4cd35ee30d96902/green_stocks.xlsm) file. 

The refractored code can be found in the [VBA_Challenge.xlsm](https://github.com/banasibb/VBA_Challenge/blob/ef0a4486eb6bdf96a4eadfe63b899f4c3d182fd8/VBA_Challenge.xlsm) file.
 <br />
![Chart 1](https://github.com/banasibb/VBA_Challenge/blob/f499aaeda4aa48ac03775ecadc7e7c3196c8f105/VBA_Challenge_2017_ss.PNG)
![Chart 1](https://github.com/banasibb/VBA_Challenge/blob/75de8f66141e0fb4ba3330ad0ce864a1afa1f611/VBA_Challenge_2018_ss.PNG)<br />

## Summary
The improvements to the refractored code included an approximately 85% improvement in runtime, as both macros executed in approximately one-sixth of the original time. Refractoring code can particularly important when working with very large data sets, as they can require a significant amount of time to run.


