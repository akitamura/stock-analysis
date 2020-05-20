# stock-analysis
Used for Bootcamp Module 2: VBA
# Challenge

### Description of new analysis:

In the new analysis it utilizes several portions of the original code. It uses the same initial and final formatting, it keeps the same tickers array, keeps the same Row Counting function, keeps the same method for adjusting to a new ticker and has the same idea for outputting the final values into cells. What was changed from the initial is the function where it goes through the list once and pulls out all the information it needs and stores it into three new arrays instead of running the loop over the data source for each ticker. This results in less overall calculations for the computer to perform. This is done by having each volume added into the tickerVolume array until a new ticker occurs, in which the new volumes would then be moved into the next array line. Before the ticker it is compared to the previous cell, it they are not equal, the starting price is recorded. If the cell after the ticker is different, then the ending price is recorded.
