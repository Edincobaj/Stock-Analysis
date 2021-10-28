# Stock-Analysis

## Project Overview

### The objective of this analysis is to refactor our stock analysis code to be able to handle more than a few stocks quickly.

## Results

### From our analysis we can see that all companies besides TERP had a positive return in 2017, this was not the case in 2018. In 2018 only 2 companies had a positive return, ENPH and RUN, while all other companies had negative returns. If we were to advise on which company to invest in it would be ENPH. This is because they had a higher overall return over the two year period.
### This can be shown in the pictures below:
![image](https://user-images.githubusercontent.com/41974323/139176248-04b2f406-e726-4de7-b2fa-1e5d861b0800.png)
![image](https://user-images.githubusercontent.com/41974323/139176526-29fa6db2-d57a-4747-8ce6-bcab300abf5f.png)
### We also see a noticable decrease in execution times when comparing the original script with the refactored script.
### The original script was taking a little over 1.7 seconds to run while the refactored script is taking around .17 seconds. This is a vast improvement in run time.
### We can atribute this to adding a ticker index along with its output arrays and readjusting our code to really make the code efficient.
> Dim tickerIndex As Single
>    tickerIndex = 0
>       Dim tickerVolumes(12) As Long
>       Dim tickerStartingPrices(12) As Single
>       Dim tickerEndingPrices(12) As Single

## Summary

### There are many advantages to refactoring code, as we can see in this analysis it can drastically reduce run times for code. Refactoring can also make code more streamlined and dynamic to any new challenges that might arise. The disadvantages are refactoring is definitely time/budget constraints, it's important to understand a projects budget and whether it is worth spending the time if there isn't a large advantage.
### The obvious advantage of our refactored code is the improvement of run time. The refactored code can also handle more data then the original code. The disadvantages would be that is that if Steve ever wanted to go in and make some changes to the code he may run into trouble if he isn't as VBA-proficient.

