# Stock Analysis with VBA (Refactored)

## Overview of Project

**Purpose**:
To refactor our existing VBA code to analyze stock volume and returns more efficiently.

[VBA_Challenge](Resources/VBA_Challenge.xlsm)

## Analysis
The addition of arrays to our code substantially decreased the time it takes for the code to run. For example, my original code ran at about 0.27s for both 2017 and 2018 worksheets. However, my refactored code ran at approximately 0.06s for 2017 and 2018 worksheets.

---
![2017](Resources/VBA_Challenge_2017.png)
![2018](Resources/VBA_Challenge_2018.png)
---

In the refactored code, the values for tickerVolumes, tickerStartingPrices, and tickerEndingPrices are generated as arrays. This allows all of the subsequent data to be populated into the array as the code runs.
'''VBA:
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Double
    
    Dim tickerEndingPrices(12) As Double
'''


In addition, instead of using a "For Loop" to iterate through the corresponding tickers in the ticker array, I created a variable called "tickerIndex" and set its value to 0. At the end of the code, the 1 is added to the value of tickerIndex. I then passed tickerIndex as an argument into each of the arrays used in this analysis.

'''VBA:
tickerIndex = 0

...

tickerIndex = tickerIndex + 1


A beneficial component of the refactored code is that it sets the value of ticker volume of all tickers to 0 upfront.
'''VBA: 
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    '''
In my original code, I did not give tickerVolume its own "For Loop", which meant that the entire code would run prior to setting the volume for the next ticker to 0. 

## Challenges
I found this assignment to be very challenging. This was in large part due to a lack of understanding on how to use arrays, especially in more complex cases. The previous sections of the modules and our discussions of arrays during class only briefly discussed arrays. It would have been helpful to have a more thorough discussion of how to implement these.

## Summary
Refactoring code can provide substantial increases to the speed that programs run. Thus, refactoring code can be important when programs will run massive datasets that might otherwise be difficult to process. However, this must be balanced by the actual time it takes for an individual to go through the refactoring process. If you know that your program runs at 0.5s already and you don't expect the dataset size to increase, then refactoring is likely not worth the time it takes to write the actual code.

While we only analyzed the total volume and return for 12 stocks, there are several thousands of stocks on the New York Stock Exchange alone. Therefore, our refactored code would be highly beneficial if we wanted to analyze a larger number of stocks. One limitation of the refactored code is that we need to manually adjust the index number to reflect the number of stocks we want to analzye. It might be beneficial to write code that determines the index variable based upon the number of tickers identified in our worksheet.
