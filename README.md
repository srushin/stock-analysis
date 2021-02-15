# VBA Stock Analysis Challenge 

## Overview

The purpose of this project was to refactor code to analyze stock data from any year within our workbook. Originally, we used nested loop statements to analyze a small set of given data. Nested loops worked for our small data set but was not as efficient as it could be. With this in mind, the goal was to refactor our code to increase the efficiency by looping through the data set once. 

## Results 

### Original Code 

My original code used nested loops to find stock changes for a specified year in my data. I created an array of the ticker values I wanted the information on. Then I used them as an index for my first (outer loop), while the inner loop went through the data looking for information that matched the indexed ticker. This returned the information I was looking for in 0.5625 and 0.578125 seconds for 2017 and 2018 respectively. The concern with this method was the length of time it took to process a years’ worth of information for only 12 stocks. 

### Refactored Code

In the pursuit of a more efficient method, nested loops were dropped in favor of several arrays and a series of loops. Here 3 different arrays were made to store the information derived by the first loop. Then a second loop displayed the information on my “VBA Challenge” sheet, by looping through the arrays. The last loop colored the Returns column using conditional statements. This allowed the code to run through the information 1 time, eliminating the extra searching of the first code, and return the same information in 0.15625 for either year. A 72.6% improvement on average!

## Summary  

Refactoring code was a great way to explore finding alternative methods to a previously successful one. It also allowed further opportunities to debug different types of coding issues. Refactoring does have disadvantages as well. If the code was not properly annotated then it could be hard to decipher the purpose of a certain line or section of code. There is also the issue of the code not working once reworked. In this project, I found that the new code did not work right away for the 2017 page as there was some problems with the data itself. While it worked in the end, it did take some time to find the cause of the issue. The data corruption was not an issue with the original code, and for our small data set, the original code did the job just fine. Having worked with both, I do like the speed and efficiency of the refactored code, but knowing its more temperamental with respect to data corruption, It may require some more refactoring to work through corruptions. 

