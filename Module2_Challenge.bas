Attribute VB_Name = "Module2_Challenge"
Sub Module2_Data()

Dim ws As Worksheet
Dim firstSheet As Worksheet

Set firstSheet = ThisWorkbook.Worksheets(1)
    Application.ScreenUpdating = False
    
For Each ws In ThisWorkbook.Worksheets
ws.Activate

    Range("I1", Range("I1").End(xlToRight)).Font.Bold = True
    Range("N1:P4").Font.Bold = True
    firstSheet.Activate
    Application.ScreenUpdating = True
    
''''''''''''''''''''''''''
'''''Column Creation '''''
''''''''''''''''''''''''''
   ws.Range("I1").Value = "Ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percent Change"
   ws.Range("L1").Value = "Total Stock Volume"
   ws.Range("P1").Value = "Ticker"
   ws.Range("Q1").Value = "Value"
   ws.Range("O2").Value = "Greatest % Increase"
   ws.Range("O3").Value = "Greatest % Decrease"
   ws.Range("O4").Value = "Greatest Total Volume"

    Dim TickerName As String
    Dim LastRowA As Long
    Dim LastRowK As Long
    Dim PreviousAmount As Long
    Dim LastRowValue As Long
    Dim SummaryTableRow As Long
    
    Dim TotalTickerVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    
    SummaryTableRow = 2
    PreviousAmount = 2
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    TotalTickerVolume = 0
    
''''''''''''''''''''''''''''
'''''Retrieval of Data '''''
''''''''''''''''''''''''''''

    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To LastRowA
        TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              TickerName = ws.Cells(i, 1).Value
                 ws.Range("I" & SummaryTableRow).Value = TickerName
                  ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
        TotalTickerVolume = 0
            
            OpenPrice = ws.Range("C" & PreviousAmount)
            ClosePrice = ws.Range("F" & i)
          
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
            ws.Range("J" & SummaryTableRow).NumberFormat = "$0.00"
                
                If OpenPrice = 0 Then
                PercentChange = 0
               
                Else
                YearlyOpen = ws.Range("C" & PreviousAmount)
                PercentChange = YearlyChange / OpenPrice
                        
            End If
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            
'''''''''''''''''''''''''
''Conditional Formting ''
'''''''''''''''''''''''''

            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
                End If
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                
            End If
        
        Next i
        
''''''''''''''''''''''''''
''''Calculated Values ''''
''''''''''''''''''''''''''
    LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To LastRowK
    
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value

            End If
            
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value

            End If
            
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

        Next i
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
    Next ws
      
    'Display a message when the summary tables are ready in every datasheet.
    MsgBox ("Your data is ready!")
    
End Sub

Sub Button2_Click()
 'Declare "ws" as Worksheet
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    Dim firstRow As Long
    Dim LastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
 
    firstRow = 1
    LastRow = 91
    firstColumn = 9 'Column I
    lastColumn = 17   'Column Q
 
    ws.Range(ws.Cells(firstRow, firstColumn), ws.Cells(LastRow, lastColumn)).Clear
 
    Next ws
End Sub
