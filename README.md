# Stock Analysis

## Overview of Project
### Background
Steve, a finance graduate, has been tasked with helping his parents look at green energy stocks in order to inform their investment choices. With the goal of diversifying their funds, he has created an excel file with data on various green stocks. He has requested that we use VBA to automate the data analysis so he can reuse these macros on new data in the future. 

### Purpose
The macros have been made to find the total number of traded shares and yearly returns for each of the green stocks for which Steve has gathered data. There is data for both 2017 and 2018, so our macro needs to receive input from the user on which year to run its analysis. The original script performed these tasks by looping through the data repeatedly for each stock. In order to reduce the run time, this script was refactored to loop through the dataset only once while gathering and storing the data for each stock. 

## Results
### Stock Performance
Trading volume can indicate momentum in a stock. As trading volume increases, this suggests a greater strength in the trend indicated by positive or negative returns. Annual returns are a measure of how much an investment has increased or decreased over the year. In other words, it indicates how much you would have lost or earned if you had purchased stock at the beginning of the year and sold on the last trading day of the year. It is clear that the selected green stocks typically performed better in 2017 than in 2018. As you can see in the tables below, in 2017 every stock except TERP had a postive return while in 2018 all but ENPH and RUN ended the year with a loss.


![2017 returns](/Resources/VBA_Challenge_2017.PNG)
![2018 returns](/Resources/VBA_Challenge_2018.PNG)

With regards to overall performance, RUN and ENPH come out as distinct front runners. Both stocks were the only two to show positive returns in both 2017 and 2018 and also showed the highest increase in trading volume, indicating strong momentum.  However, when taking trends from 2017 to 2018 (as shown below), RUN outperforms ENPH. While ENPH brought in a profit both years, its returns took a substantial dip (-0.47) in 2018 compared to 2017, while RUN came out even stronger in 2018 (+0.78).  (as calculated by 2018's total volume - 2017's total volume and displayed below). 

![Trend table](/Screenshots/Trends.PNG)

Compared to the rest, DQ performed the best in 2017 alone; however, it took a sharp turn in 2018 and pulled in the group's lowest negative return after experiencing the most significant drop. An increased trading volume suggests a solid momentum for this downwards trend. It is unlikely that Steve will recommend DQ to his parents, given the high volatility and negative trend. 
 

### Original VS. Refactored Script
The first thing I needed to do in both the original and refactored script was use an input field to ask Steve which year he wanted to look at. I used `yearValue = InputBox("What year would you like to run analysis on?")` to gather and store his response in a variable I could then use to move to the necessary excel sheet using `    Worksheets(yearValue).Activate`.  I then populated my analysis sheet with headers to describe the aggregate data that was going to be calculated for the stock activity during the year in question using: 
```vba
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
```
	
The original script's execution was then carried out by nesting the entire process of calculating each measure in a for-loop that repeated 12 times (for each of the stocks). This was done by creating a ticker array (seen below) that contained the labels for each of the stocks. This allowed the for loop to iterate through each ticker and repeat the process of locating and processing the data relevant to the stock in question. 
```vba
    Dim tickers(12) As String 
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR" 
```

Next, I simply created a variable named `ticker` that I initialized as the value within whatever index of the `tickers()` array we were in and used that as the reference value to compare the data to find the data relevant to the stock at hand.  I then nested in another for loop that ran through every row of data and, using three if statements, I added up the total trading volume of the stock `totalVolume = totalVolume + Cells(j, 8).Value`, found the starting price of the stock  `If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then  startingPrice = Cells(j, 6).Value` and found the ending price of the stock `If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then endingPrice = Cells(j, 6).Value`.  

After this I used `Worksheets("AllStocksAnalysis").Activate` to return to my analysis sheet and input my analysis using:
```vba
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
```

For the refactored code, my goal was to reduce the script's run time, so I only wanted to loop through the data set once. I used the same process for gathering the user input on year and populated the analysis headers as I did in the original script since this only needed to be done once and didn't affect run time. I also used the same ticker array to store each stock name in one of the 12 indexes. At this point, however, I deviated from the original script. I created three output arrays to store the total volume, starting price and ending price of each ticker and initialized each value in the volume array to zero using the following code: 
```vba
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
       
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
```
Next, I created a for loop that would iterate through every row of the selected data set within which the rest of the process would be nested. My code for summating the total volume and for finding and storing the starting and ending prices was the same as the original code except that I was now comparing the value in cells `Cells(i, 1)` to the value of `tickers(tickerIndex)`  now that I have the ticker index variable keeping track of which stock I am looking at. Then all I needed to do was use an if-statement to increase the value of my `tickerIndex` after I found the last value of the current ticker (as indicated by located the ending price using the `If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then` code already in place ).

Since I now had the values for each ticker and its total volume, ending price and starting price stored in my three arrays, I used a for loop to populate my analysis sheet with my findings using the code: 
```vba
For i = 0 To 11     
    Worksheets("AllStocksAnalysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = (tickerEndingPrices(i) /     	tickerStartingPrices(i)) - 1        
Next i
```

Both my original code and my refactored code were formatted using the same code and a for loop was used to conditionally format returns as red or green depending on if they were negative or positive. 

For both cases, I completed the script with a message box that displayed the total run time of my script using the timer function to take a measure of the time at the start (`startTime = Timer`) and end (`endTime = Timer`) of the script running then subtracting the end time from the start time to provide a total run time for the code. I did so using the following code. 
`MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)`

## Summary

- What are the advantages or disadvantages of refactoring code?

The clear advantage of refactoring code is that it reduces the run-time dramatically. This would certainly be very valuable as data sets grew larger since the length of time needed to run the code would increase considerably. This, in turn, would allow you to apply the same code to a much larger variety of future projects. Alongside a faster running program, this optimized code also uses memory more efficiently and consumes less of a computer's resources.

If I had to mention a disadvantage to refactoring code, I would have to say that it's slightly more complex since you need to create additional arrays and variables in order to make things more efficient. Since I am new to programming, this seems to increase the likelihood of letting simple errors slip in and makes debugging take longer. The added complexity of the overall code also makes it harder to keep track of what you are doing, as it's harder to make a mental flow chart of the task. That being said, I believe that the ability to reuse more efficient code dramatically reduces the time spent writing overall. Hence, it's well worth the initial effort to get your script as efficient as possible.

- How do these pros and cons apply to refactoring the original VBA script?

In the case of the VBA script, refactoring my code took the run-time down from 0.7347 seconds on average (N=10) to 0.1492 seconds (N=10) for the 2017 data set and down from 0.7486 seconds on average (N=10) to 0.1481 seconds for the 2018 data set. 
![Comparison of averages](/Screenshots/runtime_averages.PNG)
This is a great improvement.

With regards to the con of added complexity, I found that it didn't have too great an impact on refactoring the VBA script because I was encountering all of the elements of code for the first time while writing the original script. Things were a lot easier to make sense of as I refactored because I had already learned how to use arrays, for-loops, and conditional statements. I actually found the refactoring quite simple this time, but I could imagine that it can be an incredibly frustrating process with more complex code. I still think it will be worth it, though; this was a valuable lesson. 





