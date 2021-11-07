# stock-analysis
VBA Module , for loops , conditionals , variable etc
 ## git@github.com:Judat/stock-analysis.git
Sub Macrocheck()

    Dim testmessage As String
    testmessage = "Hello World !"
    MsgBox (testmessage)
    
End Sub
Sub DQAnalysis()

'activate worksheet
    Worksheets("DQ Analysis").Activate
'adding analsys sheet header
Range("A1").Value = "DAQO: Ticker(DQ)"

'adding row headers Ticker , Total Daily Volume and Rturn ( starting vale / ending value -1)
Range("A3").Value = "Ticker"
Range("B3").Value = "Total Daily Volume"
Range("C3").Value = "Return"


'giving values to variablkes which are needed in the analsysis( totl volkume will be the volume every timne DQ comes in loop,
totalVolume = 0
rowstart = 2
RowCount = 3113

Dim ticker As String
ticker = "DQ"
Dim startingPrice As Double
Dim endingPrice As Double




Worksheets("2018").Activate
'how is the loop going to to run , i is goin to start from zero till the end i.e. last row, any time DQ comes firsrt column in the loop,should add the value in column 8

For i = rowstart To RowCount
    If Cells(i, 1).Value = ticker Then
    totalVolume = totalVolume + Cells(i, 8).Value
    End If
    
    If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
    startingPrice = Cells(i, 6)
    End If
    
    If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
    endingPrice = Cells(i, 6).Value
    End If

Next i
MsgBox (totalVolume)
MsgBox (startingPrice)
MsgBox (endingPrice)

Worksheets("DQ Analysis").Activate

Range("A4").Value = ticker
Range("B4").Value = totalVolume
Range("C4").Value = (endingPrice / startingPrice) - 1

End Sub


Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (2018)"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(11) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   
   '3a) Initialize variables for starting price and ending price and totalVolume
   Dim startingPrice As Single
   Dim endingPrice As Single
   totalVolume = 0
   '3b) Activate data worksheet
   Worksheets("2018").Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
    
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

End Sub

Sub formatAllStocksAnalysisTable()

'formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Font.Bold = True
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous


    Range("A3:C3").Font.ColorIndex = 1

'formatting numbers
    Range("B3:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"

'using auto fit code
    Columns("B").AutoFit

'conditional formatting, of cell ibnterior using for loop

    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3).Value < 0 Then
        Cells(i, 3).Interior.Color = vbRed
     
        ElseIf Cells(i, 3).Value > 0 Then
        Cells(i, 3).Interior.Color = vbGreen
    
        Else: Cells(i, 3).Interior.Color = x1None
    
        End If

    Next i
    

End Sub

' chess board

Sub chess()

    For i = 3 To 10
    For j = 8 To 15

    Dim chessboard As Integer

    chessboard = (i + j)

        If (chessboard Mod 2 = 0) Then
        Cells(i, j).Interior.Color = vbBlack
        Else: Cells(i, j).Interior.Color = vbWhite
        End If
        
    Next j
    Next i
    
End Sub
