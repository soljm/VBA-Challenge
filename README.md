# VBA-Challenge
From Module 2: VBA Scripting from the Data Analytics Boot Camp by Monash University.

By implementing skills learnt throughout the module, an attempt at the challenge has been submitted here.

## Contents
- Screenshots of the results for the years: 2018, 2019, 2020
- VBA Script of the challenge

## Explanations
### Main Challenge
The script was designed to loop through the rows of data to summarise the data for each ticker.

A `For Loop` and `If Loop` was used to find if the tickers matched the row before it, where if the ticker matched, the total stock volume would add up to the total of stock volume for that specific ticker. If the ticker did not match, then current stock volume would be added to the summary table along with the ticker and the stock volume resets for the next ticker. 

The code for section of the challenge was taken from the solved *Credit Card Checker* student activity during the third class for this module and adjusted accordingly to the challenge.

In this `For Loop`, the yearly changes and percent changes are also calculated. The open value of the ticker was initially set for the first ticker and then changed according to the ticker changing. The close value was set as it would be the last row for a specific ticker. Then yearly change was calculated by subtracting open value from close value. 

Percent change was also calculated by dividing the yearly change with the open value and then formatting the column into percentage. The following conditional formatting was applied to percent change:
- Green cell colour for positive change
- Red cell colour for negative change

### Bonus Challenge
An initial value was set for the variables for ticker and value. An `If Loop` was used to determine whether the values of the following row was greater than or lesser than the variable value and changed accordingly. Then the value was printed their respective cells.
