# Stock Analysis Refactor

## Overview

In a previous contract I had assisted my client to analyze several green stocks (environmentally friendly investments) with the use of a VBA macro for Excel. This was very managable and fruitful, and my client appreciated the perspective my analysis gave. So they have requested a broadend scope in order to use my macro or an analysis of the entire stock market. I will refactor my previous code to work more efficiently on a larger scale.

## Results

### Stock Analysis

One thing became imedietly clear while comaring the 2017 green stocks to the 2018 performance. In 2017 only one stock lost value, but in 2018 only two stocks had gains. I will advise my client to invest in ENPH and RUN, as they were the only companies that managed to sustain their growth. 

### Code Refactoring

I was successfully able to increase the efficiancy of the macro by around 20%. This was accomplished by removing a nested "for loop" and completing the same tasks in one loop. Full comparison here:

![New Code/Old Code](https://github.com/Olibabba/Week2_Excel_HW/blob/main/resources/Screen%20Shot%202022-04-09%20at%2010.12.54%20PM.png)

I created two new arrays, tickerIndex() and tickerVolume(), and used these to remove a nested loop. Part of this process meant running a quick loop in initialize all of the tickerVolumne() values to 0:
```
For j = 0 To 11
        tickerVolume(j) = 0
Next j
```

This also allowed me to remove the results output from the main loop and put it into an isolated loop. I was also able to add the formatting into this macro, while it was seperate previously. Even with this additional task the time to complete was reduced significantly, as seen below.

From my tests:

- Old 2017 Macro -- .6015625 seconds
- New 2017 Macro -- .109375 seconds
- Improvement of 18%

![2017 Macro Performance](https://github.com/Olibabba/Week2_Excel_HW/blob/main/resources/VBA_Challenge_2017.png)

- Old 2018 Macro -- .6015625
- New 2018 Macro -- .125
- Improvement of 21%
![2018 Macro Performance](https://github.com/Olibabba/Week2_Excel_HW/blob/main/resources/VBA_Challenge_2018.png)



## Summary

Overall the refactoring was certainly succesful. The time saved on this small scale wil pay dividends when ran on a larger data set. This will save my client time in the long run which is always a good thing. The down side of this is that the time saved is on a small scale. Seconds will be saved every week, less than a minute will be saved over the course of a month. 

While refactoring the original code did improve it, the refactored code faces the same limitations as the original. Namely, if the data set is not organized by ticker, in a sequential manner, this code will be useless.

So while I am pleased with the results of the refactoring, this time may have been better spent making a more functional macro.