Sub Summerize()
    
For Each ws In Worksheets

' Variables
    
        Dim LastRow As Long
        LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        
        Dim Ticker_Name As String
        
        Dim Ticker_Total As Double
        Ticker_Total = 0
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim Diff_Price As Double
        Diff_Price = 0
        Dim End_Price As Double
        End_Price = 0
        
        Dim Next_Row As Double
        Next_Row = 0
        Dim Start_Row_Price As Double
        Start_Row_Price = 0
        
        Dim Close_Price As Double
        Close_Price = 0
        Dim Open_Price As Double
        Open_Price = 0
        
        Dim Orig_Open_Price As Double
        Orig_Open_Price = 0
        Dim Change_Price As Double
        Change_Price = 0
        
        Dim Change_Percent As Double
        Change_Percent = 0
    
        
        
'Add the word Ticker to the First Column Header
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Volume"
             
'Add Values to Summary
    
    For i = 2 To LastRow
        
        
    'Check if we are still within the same ticker, if it is not...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'Set the name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the  Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
      
      Open_Price = ws.Cells(i, 3).Value
      
      Close_Price = ws.Cells(i, 6).Value
      
      Orig_Open_Price = Open_Price - End_Price
      
      Change_Price = Close_Price - Orig_Open_Price
      
      Change_Percent = 1 - (Close_Price / Orig_Open_Price)
      
      ' Print the name in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
      
      ' Print the change amount in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Change_Price
      
        ' Print the change percentage in the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Change_Percent

      ' Print the Amount to the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total
      Ticker_Total = 0
      End_Price = 0

    'If the cell immediately following a row is the same ..
    
    Else
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        Start_Row = ws.Cells(i, 3).Value
        Next_Row = ws.Cells(i + 1, 3).Value
        Diff_Price = Next_Row - Start_Row
        End_Price = End_Price + Diff_Price
    End If
  Next i
     
        
' condtional formatting1


     For r = 2 To LastRow
     If ws.Cells(r, 11).Value > 0 Then
                ws.Cells(r, 11).Interior.ColorIndex = 4 'green
            Else
                ws.Cells(r, 11).Interior.ColorIndex = 3 ' Red
            End If
       Next r
       
' condtional formatting2
             
   For p = 2 To LastRow
     If ws.Cells(p, 12).Value > 0 Then
                ws.Cells(p, 12).Interior.ColorIndex = 4 'green
            Else
                ws.Cells(p, 12).Interior.ColorIndex = 3 ' Red
            End If
       Next p
       
' min and max of summary table

    Dim High_Change As Double
    High_Change = 0
    Dim High_Ticker As String
    
    Dim Low_Change As Double
    Low_Change = 0
    Dim Low_Ticker As String
    
    Dim High_Value As Double
    High_Value = 0
    Dim HighV_Ticker As String



'Add the word Ticker to the First Column Header
        ws.Range("P2").Value = "Greatest Increase"
        ws.Range("P3").Value = "Greatest Decrease"
        ws.Range("P4").Value = "Greatest Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
         
        
    For v = 2 To LastRow
    If ws.Cells(v, 12).Value > High_Change Then
        High_Change = ws.Cells(v, 12).Value
        High_Ticker = ws.Cells(v, 10).Value
    ' Print the summary 2 table
        ws.Range("Q2").Value = High_Ticker
        ws.Range("R2").Value = ws.Cells(v, 12).Value
    Else
    End If
    Next v
    
   For l = 2 To LastRow
    If ws.Cells(l, 12).Value < Low_Change Then
        Low_Change = ws.Cells(l, 12).Value
        Low_Ticker = ws.Cells(l, 10).Value
    ' Print the summary 2 table
        ws.Range("Q3").Value = Low_Ticker
        ws.Range("R3").Value = ws.Cells(l, 12).Value
    Else
    End If
    Next l
    
    For c = 2 To LastRow
    If ws.Cells(c, 13).Value > High_Value Then
        High_Value = ws.Cells(c, 13).Value
        HighV_Ticker = ws.Cells(c, 10).Value
    ' Print the summary 2 table
        ws.Range("Q4").Value = HighV_Ticker
        ws.Range("R4").Value = ws.Cells(c, 13).Value
    Else
    End If
    Next c
    
 Next ws

End Sub