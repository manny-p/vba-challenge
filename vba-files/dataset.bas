Attribute VB_Name = "Module1"

Sub StockData()

' Declare variables

    Dim i As Integer
    Dim j As Long
    Dim k As Integer
    Dim Ticker As String
    Dim OpenPrice As Currency
    Dim Closeprice As Currency
    Dim TotalSheets As Integer
    Dim TotalVolume As LongLong
    Dim MinPercentChange As Double
    Dim MaxPercentChange As Double
    Dim MaxTotalVolume As LongLong
    Dim TickerMinPercent As String
    Dim TickerMaxPercent As String
    Dim TickerMaxVolume As String
    Dim OutputRow As Integer
    Dim LastRow As Long
    
    
' Assign worksheets to TotalSheets
   
TotalSheets = ActiveWorkbook.Worksheets.Count

' Loop through sheets 

For i = 1 To TotalSheets

    Worksheets(i).Activate
    OutputRow = 2
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
' Column output 

    Cells(1, 9) = "Ticker Symbol"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
' Bonus output
    
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    
    TotalVolume = CLng(0)
                
    For j = 2 To LastRow
    
        Ticker = Cells(j, 1)
        TotalVolume = TotalVolume + CLng(Cells(j, 7))
        
        If Cells((j - 1), 1) <> ticker Then
        
            OpenPrice = Cells(j, 3)
            
        End If
        
        If Cells((j + 1), 1) <> ticker Then
                                 
               Closeprice = Cells(j, 6)
               Cells(OutputRow, 9) = Ticker
               Cells(OutputRow, 10) = Closeprice - OpenPrice
               Cells(OutputRow, 12) = TotalVolume
               TotalVolume = 0
                        
               If Cells(OutputRow, 10) < 0 Then
                    Cells(OutputRow, 10).Interior.ColorIndex = 3
                    Cells(OutputRow, 11).Interior.ColorIndex = 3
            
               Else
            
                    Cells(OutputRow, 10).Interior.ColorIndex = 4
                    Cells(OutputRow, 11).Interior.ColorIndex = 4
            
               End If
                       
           
            Cells(OutputRow, 11).NumberFormat = "0.00%"
            If OpenPrice <> 0 Then
            
                Cells(OutputRow, 11) = (Closeprice - OpenPrice) / OpenPrice
                
            Else
            
                Cells(OutputRow, 11) = "Not Available"
            
            End If
                               
            OutputRow = OutputRow + 1
           
        End If
           
    Next j
 
    MaxTotalVolume = Cells(2, 12)
    MaxPercentChange = Cells(2, 11)
    minpercentchange = Cells(2, 11)
    
    For k = 3 To OutputRow
    
        If Cells(k, 12) > MaxTotalVolume Then
            MaxTotalVolume = Cells(k, 12)
            TickerMaxVolume = Cells(k, 9)
        End If
        
        If Cells(k, 11) > MaxPercentChange And Cells(k, 11) <> "Not Available" Then
            MaxPercentChange = Cells(k, 11)
            TickerMaxPercent = Cells(k, 9)
        End If
        
        If Cells(k, 11) < minpercentchange And Cells(k, 11) <> "Not Available" Then
            minpercentchange = Cells(k, 11)
            TickerMinPercent = Cells(k, 9)
        End If
        
    Next k
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(2, 16) = TickerMaxPercent
    Cells(2, 17) = MaxPercentChange
    Cells(3, 16) = TickerMinPercent
    Cells(3, 17) = minpercentchange
    Cells(4, 16) = TickerMaxVolume
    Cells(4, 17) = MaxTotalVolume
    
    Worksheets(i).Columns("A:Q").AutoFit
    
Next i

VBA.Interaction.MsgBox "Booom!!!"
    
End Sub











