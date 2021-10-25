# Stock Analysis: Excel VBA
## Overview of Project
Steve wanted to understand stock market trends for a dozen stocks in 2017 and 2018. He appreciated the workbook that we prepared for him that can analyze his dataset with just one button. He would like to expand his data set to include the entire stock market (rather than just a select few). In order to efficiently deliver this information, we will have to refactor the original code. 
### Purpose
This report was created to discover information on Stock Data in 2017 and 2018 by refactoring the VBA script to determine whether or not the stocks are worth investing. Refactoring the code will make retrieving this data faster and more simple.
## Results
To begin refactoring our code, we created a `tickerIndex` (set to 0) and three output arrays: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. 
![1ab](https://user-images.githubusercontent.com/92230478/138617847-6ec97ea0-c0b9-42cb-a0f4-196bf194cb34.PNG) 

We added in a `for` loop to increase `tickerVolumes` based on the `tickerIndex`. 
![3abc](https://user-images.githubusercontent.com/92230478/138618041-18bc7da2-4540-445d-b88e-603c5b6c14a2.PNG)

Next, `if-then` statements included to establish `tickerStartingPrices` and `tickerEndingPrices` as well as to increase the `tickerIndex`. 
![4g](https://user-images.githubusercontent.com/92230478/138618038-10244172-fca6-487b-bf0f-aa19a5b28d56.PNG)

Finally, we placed a `for` loop through the arrays to output in the "Ticker," "Total Daily Volume," and "Return" columns in the spreadsheet.
![4a](https://user-images.githubusercontent.com/92230478/138617927-a59c4eba-48f8-4443-9d59-c3330cf8c571.PNG)
With the new refactored code being complete, we were able to run the Stock Analysis and were given the exact same outputs as the original code from the Module. The only difference noticed is that the refactored code has a quicker elapsed run time. The original code run time was about 0.9 seconds, while the refactored one was about 0.2 seconds (see below for photos)
### 2017 Results and Timing
![2017](https://user-images.githubusercontent.com/92230478/138617137-a587b216-4f0a-4b8e-88a7-96fe3210f079.PNG)
The stocks had a great year in 2017 (all except "TERP")! 11/12 stocks had a positive return with 4 of those having 100.0%+.
![2017 old](https://user-images.githubusercontent.com/92230478/138618427-46c1e266-13dc-43b9-a744-0affe2d517fd.PNG)
As pictured above, the original script for stock analysis using VBA had an elapsed run time of 0.875 seconds.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92230478/138618432-d02f8426-f2f4-4381-ba70-8f8219077755.png)
With the new refactored code, the elapsed run time shortened to about 0.18 seconds.

### 2018 Results and Timing
![2018](https://user-images.githubusercontent.com/92230478/138617139-266495f6-923b-4571-8fde-ee6f980fbaf2.PNG)
The stocks did not perform as well in 2018 as they did in 2017. "ENPH" and "RUN" are the only two with positive returns at 80.0%+. The other 10 stocks all have negative returns ranging from -3.5% to -62.6%.
![2018 old](https://user-images.githubusercontent.com/92230478/138618435-42227200-054e-4675-9288-b3c4b6202730.PNG)
The original script from the module for 2018 had an elapsed run time of about 0.887 seconds.
![VBA_Challenge_2018](https://user-images.githubusercontent.com/92230478/138618436-3db63d6c-1058-4a14-bffd-e6098949a25a.png)
The new refactored code shortened the elapsed run time to about 0.245 seconds.

## Summary
### Advantages and Disadvantages of Refactoring Code
Refactoring makes the code script appear cleaner as it gets straight to the point. This is beneficial for those working with significantly larger codes as it makes the appearance of bugs and other unwanted pieces more obvious. However, with editing code there are also downsides that occur. The most impactful disadvantage is unintentionally losing outcomes by adding/removing functions incorrectly. 
### Advantages and Disadvantages of Original and Refactored VBA Script
The refactored VBA script in our spreadsheet was beneficial as it decreased our elapsed run time for Stock Analysis from 0.9 seconds to 0.2 seconds by making the code easier to read. The software and I both found the refactored code to be a condensed version of the original. I did not find any disadvantages with this script, however if the script was long, it would haven been difficult to refactor it as there are more functions to keep track of.
