# stock-analysis
module 2 stock analysis project VBA

### Overview 
The purpose of this project was to help Steve prepare this stock data workbook to present to his parents and help them pick the best stock(s) to invest in. In this module we learned several programing functions in VBA to sort through the information easily and quickly. Some of the functions we learned in this module were learning how to enable the macro, create loops to pull information from one sheet and condense it and put it in another using conditionals. 
    
This workbook of stock data is showing data from 2017 and 2018 stocks in twelve green energy companies. The different tickers show what the starting price and ending price of each ticker was on that given day in the year. Steve wants to use this information to compare the prices of other stock companies to the Daqo company his parents want to invest in. 

### Results

Steve created this large data worksheet of information from all twelve companies, and we used our new VBA skills to help him pull the information from the large spreadsheet. To start we were given a bit of code to refactor to pull out what the “Total Daily Volumes” and the “Return” was of each stock. We had to pull this information from the given data which was the date of the stock, opening, closing, adjusted, highest and lowest prices of the stock and the volume. 

<img width="637" alt="VBA_challenge_loops" src="https://user-images.githubusercontent.com/45208773/131945016-d91f31ce-0bcd-4d22-9279-3d3b0a36f5e4.PNG">


We wrote a nested loop, shown above, that went through the data and gathered the ticker, starting price and ending price. From the information we obtained from this loop, we used in the next loop to gather the volumes of each stock. The last piece of code outputs all the information we collected on to a new sheet and gives us “Total Daily Volume” and “Return” 

<img width="524" alt="totalVolume_Returns" src="https://user-images.githubusercontent.com/45208773/131945149-e71f0993-9e98-4fb5-a096-29d149acf21e.PNG">

When you run the entire macro, you will first get a box that asks, “what year would you like to run the analysis on?”  You then type in the year, and it will give you that year’s total daily volume and return for each stock. It will also highlight the return red if there was a negative return in the stocks. 

<img width="326" alt="2017" src="https://user-images.githubusercontent.com/45208773/131945434-55323207-007e-46aa-8302-9001054573b4.PNG">
<img width="316" alt="2018" src="https://user-images.githubusercontent.com/45208773/131945444-1b3cb7e9-075a-403c-a8c4-6c813d013482.PNG">

You can see in the snip-its above that the year 2017 had a much better return of stocks than 2018. In 2017 Daqo had almost a 200% return versus in 2018 it was -63% return. 

We also had a bit of code in there, using the Timer function, that produces a message box of how long it took the code to run to give us all our information. 

<img width="277" alt="2017time" src="https://user-images.githubusercontent.com/45208773/131945792-256d8904-3784-4d37-904b-c6c0372eb59c.PNG">
<img width="298" alt="2018time" src="https://user-images.githubusercontent.com/45208773/131945803-93aa610a-48f4-40b1-9c48-223664c943ed.PNG">

## Advantages 

The advantage of refactoring code makes it easier for one person or multiple people to use and read. When you refactor the code, you are kind of performing a “clean up” on the already written code. This is beneficial in a group-project situation. 

## Disadvantages 

Some of the disadvantages of refactoring code would be over-refactoring to the point where it could completely change the outcome of what you originally wanted. Another disadvantage could be refactoring code when it is not needed, like a “don’t fix it if it isn’t broke” situation because then it could mess up what good code was already written. 





