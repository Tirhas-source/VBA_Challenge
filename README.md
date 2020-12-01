# VBA_Challenge

Sub stock()


Dim ws As Worksheet
Dim ticker As String
Dim LastRow As LongPtr
Dim New_Table As Integer
Dim Yearly_Change, Year_Open, Year_Close As Double
Dim open_close_sum As Double
Dim Vol As Double


New_Table = 2
open_close_sum = 2
Vol = 0


For Each ws In Worksheets
   
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    '#1
    ws.Cells(1, 8).Value = "Ticker"
    '#2
    ws.Cells(1, 9).Value = "Yearly_Change"
    '#3
    ws.Cells(1, 10).Value = "Percent_Change"
    '#4
    ws.Cells(1, 11).Value = "Total_Stock_volume"

    

    'loop for ticker
        For i = 2 To LastRow
             Year_Open = ws.Cells(open_close_sum, 3)
             ticker = ws.Cells(i, 1).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            Year_Close = ws.Cells(i, 6).Value
            Vol = ws.Cells(i, 7).Value
       
            
          

            If Year_Open = 0 Then
                ws.Cells(New_Table, 10).Value = Null
                
            Else
            ws.Cells(New_Table, 10).Value = (Year_Open - Year_Close) / Year_Open
            
              
            'data to New_table
            ws.Cells(New_Table, 8).Value = ticker
            ws.Cells(New_Table, 9).Value = Year_Open - Year_Close
            ws.Cells(New_Table, 10).Value = (Year_Open - Year_Close) / Year_Open
            ws.Cells(New_Table + 1, 11).Value = Vol
            New_Table = New_Table + 1
        
            ws.Cells(New_Table, 10).NumberFormat = "0.00%"
        
        
            If ws.Cells(New_Table, 9).Value > 0 Then
                ws.Cells(New_Table, 9).Interior.ColorIndex = 4
            Else
            ws.Cells(New_Table, 9).Interior.ColorIndex = 3
            
            End If
            
        End If
        
        
End If

    

Next i
    


Next ws


End Sub



