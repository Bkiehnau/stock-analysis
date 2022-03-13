#                               Excel VBA Stock Analysis
The purpose of this challenge was to utilize our newly learned skill to analyze the stock data provided from 2017 and 2018 into a table by ticker. Each ticker should show total daily volumes and the % of return by color code. Additionally, this should all be done using code in the VBA function instead of using formulas in the actual spreadsheet.
#                                Challenge Analysis
This challenge was very similar to the module work that was done earlier in the week, except that we were provided starter code as a basis instead of creating code from scratch and building upon it throughout the modules. From the starter code we were provided our tickers and some basic formattting. From there we were given four tasks to complete.
#                                Task 1a
Task 1a was to create a ticker index. I did this by using the following code tickerIndex = tickers(i) and then I decided to create a for statement that set up variable i. For i = 0 to 11.
#                                Task 1b
Task 1b asked us to create output arrays for three variables. I did this by creating dimensions based upon the instructions provided. These instructions asked us to make tickerVolume long, tickerStartingPrice a single and tickerEndingPrice a single. I did this  by providing the following code.
Dim tickerVolumes As Long
Dim tickerStartingPrices As Single, tickerEndingPrices As Single
#                                Task 2a
Task 2a asked us to create a for loop to initialize the tickerVolumes to zero. In order to do this I created nested loop with j. The code that I used is as follows.
tickerVolumes = 0
For j = 2 to RowCount
#                                Task 2b
Task 2b asked us to loop over all rows in the spreadsheet. I did this with an If, Then function. The following code is what I achieved. If Cells(j, 1).Value = tickerIndex Then tickerVolumes = tickerVolumes + Cells(j, 8).Value
#                                Task 3a, 3b, 3c, and 3d
Task 3 asked us to increase the volume for the current ticker, check if the current row is the first row with the selected tickerIndex, check if the current row is the last row with the selected tickerIndex and increase the tickerIndex. Up to this point the challenge had been fairly easy. This had me looking back at what we learned to review loops and functions from our classes during the week. In the end I used the following code.
If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
    tickerStartingPrices = Cells(j, 6).Value
If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
    tickerEndingPrices = Cells(j, 6).Value
#                                Task 4
Task 4 asked us to loop through our arrays to output the daily volumes and returns by ticker. To me this seemed fairly straightforward and I used the code that follows.
Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + i, 2).Value = tickerIndex
           Cells(4 + i, 3).Value = tickerVolumes
           Cells(4 + i, 4).Value = tickerEndingPrices / tickerStartingPrices - 1
#                                Challenges
The main challenge that I came across while writing code in VBA was the syntax of my code. I had to keep reminding myself to add a period or a parentheses. I would occaissionally forget quotation marks and in general reminding myself to close my functions with and end if or next i. These are all things I have tried to put into my brain bank moving forward to Python.
#                                Results
#                Compare Stock performance between 2017 and 2018.
When comparing the end results from 2018 to 2017 the returns were clearly more favorable in 2017. 2018 only had two positive returns and ten negative returns. The two positive returns were "ENPH" and "RUN" at 81.9% and 84% respectively. In turn, 2017 had eleven psoitive returns and one negative return, "TERP". In 2017 "SPWR" and "FSLR" had the highest daily volume both being above 600,000. In 2018 "SPWR" was again in the top two along with "ENPH" for daily volume. I am not sure I did my original green stocks analysis right as when I ran the macro for 2017 it was 50446.42 and for 2018 it was 50446.36. After I rafactored my code it was 1.144531 for 2017 and .640625 for 2018. These seem like vast differences and I am unsure why it would jhave changed so much. Was my code really that much more efficient? I will have to ask these questions in class. Please see attached images for reference.
#                What are advantages and disadvantages of refactoring code?
The big advantage to refactoring code is that you can end up with hgiher quality code. Through testing and tweaking code, you can get to either more effective or more efficient code. A big disadvantage is that it all depends on how good of comments you added to your good while writing it. If you did not make diligent comments it could lead to having to rewrite large portions of code or do multiple tests of functionality.
#                How do these advantages and disadvantages apply to refactoring original VBA   script?
This applies to original VBA script because you need to keep in mind refactoring code as you write your original. If you write code that is aloppy or not easy to follow, then you will not be able to go back and refactor your code as easily.
