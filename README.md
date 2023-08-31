# VBA-Challenge
# multiple year stock data

The name of the macro is "PrintResult" which in turn calls another method "InsertSummary" to actually print all the data on the sheet.

The PrintResult macro selects the range of cells which has the data and then stores the entire range as a 2 dimensional array.
Each row of the array represents the daily record for the ticker and the column represents the fields/attributes.

As we iterate through the array we would store the values in variables , once such variable "ticker" stores the current ticker name whose data is being read when a new ticker name is encountered or the end of the record is reached then the procedure "Insert Summary" is invoked with the following parameters

row -> the current row in the array 
ticker -> the name of the ticker whose summary needs to be printed.
yearlychange -> the difference of closing stock and opening stock value for the ticker
openingstock -> the openingstock value for the ticker ( the first for the month)
totalStockVol-> the sum of all the vol for the given ticker

The "InsertSummary" procedure prints the summary report on the worksheet and formats the column according to present data as required.


