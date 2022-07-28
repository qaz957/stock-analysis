# Stock Analysis With VBA

## Purpose
  This project is mean to analyze data about stocks to give a more clear look at stock trends. We'll compare the intial nested loop method to the refactored method we use here to see which code runs more efficiently. Ideally, by taking fewer steps we can make a cleaner, more efficient script.
  
## Analysis
### Nested Loop
  The first method we used to parse out the data we needed, was a nest loop method to run through the data on the Excel sheet:
  
    For i = 0 To 11

         ticker = tickers(i)
         totalVolume = 0

      'Activate data worksheet
         Worksheets(yearValue).Activate
         For j = 2 To RowCount

          'Find Total Volume
             If Cells(j, 1).Value = ticker Then

                  totalVolume = totalVolume + Cells(j, 8).Value

              End If

          'Find Starting Price
              If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  startingPrice = Cells(j, 6).Value

              End If

          'Find Ending Price
              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  endingPrice = Cells(j, 6).Value

              End If

          Next j

     'Output results
       Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i
    
 ![2017 Time Initial](https://user-images.githubusercontent.com/108296899/181590378-7375d094-8fb4-499f-ba31-92a142aef742.png)
 ![2018 Time Initial](https://user-images.githubusercontent.com/108296899/181590388-22620568-bb42-4c62-8adc-30039e378693.png)

 
 ### Refactor 
  Our second method involves using multiple arrays to store the data and then output it by pulling the stored values from the arrays after the inital loop had ran through the data using a universal index:
  
    '1a) Create a ticker Index
      tickerIndex = 0

    '1b) Create three output arrays
      Dim tickerVolumes(12) As Long
      Dim tickerStartingPrices(12) As Single
      Dim tickerEndingPrices(12) As Single
    
      For i = 0 To 11
          tickerVolumes(i) = 0
          tickerStartingPrices(i) = 0
          tickerEndingPrices(i) = 0
      Next i

    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

            '3a) Increase volume for current ticker
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                End If

            '3c) check if the current row is the last row with the selected ticker
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                   tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                End If

            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                   tickerIndex = tickerIndex + 1
                End If

        Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

    Next i 

![2017 Time Refactor](https://user-images.githubusercontent.com/108296899/181590474-3e5187d8-c3dc-4378-9251-5f0747436544.png)
![2018 Time Refactor](https://user-images.githubusercontent.com/108296899/181590484-ab019002-cbc7-470b-9fc6-540d78664cd9.png)



We can see that refactoring the script to use multiple arrays to store the data, we were able to run the code much faster. Scaled up, this can save an huge amount of time and energy when applied to much larger datasets.

## Summary
### Refactoring Code
 Refactoring code is a good habit because it will save time and energy, especially if you plan on running the scripts over large sets many times over. The drawbacks may come in the way of specialization. however, by tailoring your code run more effieciently on one specific dataset, you run the risk of making it less applicable to more universal application.
 
 Ultimately the script we created after refactoring was more effifient. As well as that, it is cleaner due to its use of arrays which allows for easier editing and debugging in the future if we have to adapt this method to another dataset. Our origninal script gives the same result, but one could argue the nested loop feature allows you apply it more broadly to other datasets.
