# Stocks Analysis

## Overview of Project
This report is a stock analysis of green energy companies using VBA. This analysis consisted of 12 green energy stock performances during 2017 and 2018. The analysis was executed using VBA to determine the total volume and return on investment. 

### Purpose
VBA was used to make the analysis more efficient and streamlined. The VBA code was refactored to keep analysis run time short and computers processing functions efficient. The refactored code also enables the new input of data from a given year into the analysis.

### Results: All Stocks Analysis Refactored 2017
The picture below is a screenshot of the 2017 analysis using the refactored VBA code. Using the timer function:    

Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub. 

Excel is able to count the duration of time it took to process the code and generate an answer. Although times vary, the refactored code completed the analysis in under one second versus the original code taking over a second. 

![VBA_Challenge_2017](Resource/VBA_Challenge_2017.png.png)

### Results: All Stocks Analysis Refactored 2018
2017 brought great returns for 9 out of the 12 stocks in the analysis. The highest return soared at 200%. 2018 was not as good to the green energy portfolio with only 2 out of the 12 stocks bringing a return on investment. The other 10 stocks go into the negative with the lowest bottom at -63%. The comparison between 2017 and 2018 can depict a major correction within green energy stocks.

![VBA_Challenge_2018.png](Resource/VBA_Challenge_2018.png.png)

## Summary
The advantages of refactoring code are that the analysis can be completed more efficiently and the code can be reused with new data. Refactored code is easier to read and enables a team to edit the code. The disadvantage of refactoring is the time it takes to rewrite a working code. If the code isn’t properly stored then lines of code can be lost and new errors can appear.

The original VBA code worked very well and only took a little over a second to execute. The refactored code took a long time to develop and worked under a second faster than the original code. The refactored code was easier to read and to build upon than the original code. The pros and cons do apply to the refactored code but it was necessary for a successful analysis.

