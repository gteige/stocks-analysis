##Purpose
Through VBA macros running complicated excel code is made simple. When Steve wanted to look at a broader selection of stocks, I was easily able to refactor, or modify existing code, to determine returns on the additional stocks. In refactoring the code in VBA, the end-user and coder experiences were improved, as the code processes ran faster and the code itself was better organized and easier to navigate.
##Background
Steve has been giving advice to family members based on the original code developed to determine returns on a selection of stocks. Steve began with data for 2017, and 2018 for 12 stocks. He is wanting to look at a wider collection of stocks, for the macro to run smoothly with a greater volume of data it needs to be refactored to allow for quick analysis. 
##Results
Through refactoring my code to analyze all stocks it becomes obvious that 2017 was a better year for the stocks in consideration with 11 of the 12 stocks experiencing positive retuns. In contrast, only 2 of the same group experienced positive returns for the year 2018. In refactoring the code, the analysis for both years is completed in a shorter period. 
###Execution Times
For data for 2017, the original code ran in .2890625 seconds. With the refactored code, the analysis of 2017 ran over three times faster in just .0703125 seconds.    
![2017]()
For 2018, the original code ran in .2578125 seconds. With the refactored code it also ran more than three times faster in .078125 seconds.
![2018]()
###Code Improvements
In refactoring the code, the three outputs are initialized as arrays, instead of just variables. 
```
    Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
 ```           
Arrays are simpler and more efficient in code, and as such the run time on both sets of data decreased a signifcant amount. Arrays are also beneficial in being reusable once declared, which would be helpful for Steve if he were to use this code to look at data of other stocks. Additionally, the reformated code is visually much easier to comprehend as each section has a header to explain its purpose. 
##Summary
###Refactoring Code
Refactoring code is a process of restructuring existing code to lessen complexities (which may cause for more issues) or help code be more readable, while maintaining the functionality of the code. The big pros of refactoring code are improved execution times and potentially cleaner code. On the other hand, the cons for refactoring are the potential for the code to break, and the cost of development time for refactoring. 
###Refactoring for VBA
The refactoring of the code to analyze the group of stocks definietly has both the pros and cons listed above. After refactoring the code, it is much cleaner in its present and is easy to follow. As a con, the code did break during refactoring, luckily it was as simple as a + and - being mixed-up, once I uncrossed my wires the refactoring was smooth sailing. 

