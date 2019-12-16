Sub FinalTicker()
' ----------------------------------------------------------------------------------------------------------------------
' Initialize a FOR LOOP to loop through all of the worksheets in the workbook
' ----------------------------------------------------------------------------------------------------------------------
For Each ws In Worksheets
    
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Define all variables to be used in this VBA code
  ' ----------------------------------------------------------------------------------------------------------------------
    Dim lRow As Long                                    ' last row in the active worksheet
    Dim lCol As Long                                    ' last column in the active worksheet
    Dim lColAlpha As String                             ' alpha value of the last column in the active worksheet
    Dim fCell As String                                 ' first cell in the active worksheet
    Dim lCell As String                                 ' last cell in the active worksheet
    Dim wsRange As String                               ' range of cells in the active worksheet
    Dim Ticker As String                                ' Holds the ticker symbol value
    Dim Symbol As Integer                               ' Holds the number of the ticker symbol value
    Dim iRow As Integer                                 ' INT value for the row output location when summing Total Stock Volume
    Dim iCol As Integer                                 ' INT value for the column output location when summing Total Stock Volume
    Dim OpenPrice As Double                             ' Holds the open price for a ticker symbol value
    Dim ClosePrice As Double                            ' Holds the close price for a ticker symbol value
    Dim TotVol As Long                                  ' Holds the summed Total Volume total for a ticker symbol value
    Dim GrtTick As String                               ' Ticker value for the "CHALLENGES" section
    Dim GrtPer As Double                                ' Percent change value for the "CHALLENGES" section
    Dim GrtVol As Single                                ' Volume change value for the "CHALLENGES" section
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Set initial values for cell location variables
  ' ----------------------------------------------------------------------------------------------------------------------
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row            ' This defines the initial last row in the active worksheet
    lCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column  ' This defines the initial last column in the active worksheet
    lColAlpha = Split(ws.Cells(1, lCol).Address, "$")(1)    ' This defines the initial alpha value of the last column in the active worksheet
    fCell = "A1"                                            ' This defines the initial first cell in the active worksheet
    lCell = lColAlpha & lRow                                ' This defines the initial last cell in the active worksheet
    wsRange = fCell & ":" & lCell                           ' This holds the range of initial cells in the active worksheet
    
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Clause to sort the data within the worksheet on <ticker> then on <date>
  ' ----------------------------------------------------------------------------------------------------------------------
    With ws.Sort
     .SortFields.Clear
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range(wsRange)
     .Header = xlYes
     .Apply
    End With
        
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Populate the data headers for our output fields
  ' ----------------------------------------------------------------------------------------------------------------------
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Yearly Volume"
  ' ----------------------------------------------------------------------------------------------------------------------
  ' This LOOP section defines the data that will flow into the summary table in the worksheet.
  ' ----------------------------------------------------------------------------------------------------------------------
    Symbol = 2
    For i = 2 To lRow
    
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker = ws.Cells(i, 1).Value
      ws.Cells(Symbol, 9).Value = Ticker                   ' Output the unique Ticker value from the loop into the next row of column 9
      OpenPrice = ws.Cells(i, 3).Value                     ' Store the opening price for that ticker value into memory as OpenPrice
      Symbol = Symbol + 1
      Ticker = 0
    End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ClosePrice = ws.Cells(i, 6).Value
      If OpenPrice <> 0 Then
        ws.Cells(Symbol - 1, 10).Value = (ClosePrice - OpenPrice)
        ws.Cells(Symbol - 1, 11).Value = FormatPercent(((ClosePrice - OpenPrice) / OpenPrice), 4)
      Else:
        ws.Cells(Symbol - 1, 10).Value = 0
        ws.Cells(Symbol - 1, 10).Value = 0
      End If
    End If
    
    If (ws.Cells(Symbol - 1, 10).Value < 0) Then
       ws.Cells(Symbol - 1, 10).Interior.ColorIndex = 3
    Else
       ws.Cells(Symbol - 1, 10).Interior.ColorIndex = 4
    End If
   Next i
  ' ----------------------------------------------------------------------------------------------------------------------
  ' This loop uses the Application.WorksheetFunction.SumIf function to find the Total Stock Volume
  ' ----------------------------------------------------------------------------------------------------------------------
    iCol = 9
    iRow = 2

    For j = 2 To Symbol - 1
      Ticker = ws.Cells(iRow, iCol)
      ws.Cells(iRow, iCol + 3).Value = Application.WorksheetFunction.SumIf(ws.Range("A2:A" & lRow), ws.Cells(iRow, iCol), ws.Range("G2:G" & lRow))
      iRow = iRow + 1
    Next j
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Reset values for cell location variables to pertain to the new table of summarized values
  ' ----------------------------------------------------------------------------------------------------------------------
    lRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    lCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    lColAlpha = Split(ws.Cells(1, lCol).Address, "$")(1)
    fCell = "I1"
    lCell = lColAlpha & lRow
    wsRange = fCell & ":" & lCell

  ' ----------------------------------------------------------------------------------------------------------------------
  ' This re-defines the last row in the worksheet to that of the new summary table
  ' ----------------------------------------------------------------------------------------------------------------------
    lRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Populate the data headers for our output fields for the challenge table
  ' ----------------------------------------------------------------------------------------------------------------------
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Output the "Greatest % increase" ticker symbol and value
  ' ----------------------------------------------------------------------------------------------------------------------
    GrtTick = 0
    GrtPer = 0
    
    For i = 2 To lRow
      If ws.Cells(i, 11).Value > GrtPer Then
        GrtTick = ws.Cells(i, 9).Value
        GrtPer = ws.Cells(i, 11).Value
      End If
        ws.Cells(2, 16).Value = GrtTick
        ws.Cells(2, 17).Value = GrtPer
    Next i
    
    ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17).Value)
    
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Output the "Greatest % Decrease" ticker symbol and value
  ' ----------------------------------------------------------------------------------------------------------------------
    GrtTick = 0
    GrtPer = 0
    
    For i = 2 To lRow
      If ws.Cells(i, 11).Value < GrtPer Then
        GrtTick = ws.Cells(i, 9).Value
        GrtPer = ws.Cells(i, 11).Value
      End If
        ws.Cells(3, 16).Value = GrtTick
        ws.Cells(3, 17).Value = GrtPer
    Next i
    
    ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17).Value)

  ' ----------------------------------------------------------------------------------------------------------------------
  ' Output the "Greatest total volume" ticker symbol and amount
  ' ----------------------------------------------------------------------------------------------------------------------
    GrtTick = 0
    GrtVol = 0
    GrtPer = 0
    
    For i = 2 To lRow
      If ws.Cells(i, 12).Value > GrtVol Then
        GrtTick = ws.Cells(i, 9).Value
        GrtVol = ws.Cells(i, 12).Value
      End If
        ws.Cells(4, 16).Value = GrtTick
        ws.Cells(4, 17).Value = GrtVol
    Next i
    
    
  ' ----------------------------------------------------------------------------------------------------------------------
  ' Move to the next worksheet. If no further worksheets, then print a MsgBox to inform user that update is complete
  ' ----------------------------------------------------------------------------------------------------------------------
    
    Next ws

    MsgBox ("Update Complete")

End Sub
