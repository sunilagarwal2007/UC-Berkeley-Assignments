Sub Analyze_Stock_Market_Data()
'Create a script that will loop through one year of stock data for each run
'and return the total volume each stock had over that year.
'You will also need to display the ticker symbol to coincide with the total stock volume.
'Your result should look as follows (note: all solution images are for 2015 data).

Sheets.Add.Name = "Combined_Data"
Sheets("Combined_Data").Move Before:=Sheets(1)
Set combined_sheet = Worksheets("Combined_Data")

' Loop through all sheets
   For Each ws In Worksheets
       ' Find the last row of the combined sheet after each paste
       ' Add 1 to get first empty row
       lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
       ' Find the last row of each worksheet
       ' Subtract one to return the number of rows without header
       lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
       ' Copy the contents of each state sheet into the combined sheet
       combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
   Next ws
   
    ' Copy the headers from sheet 1
   combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
   ' Autofit to display data
   combined_sheet.Columns("A:G").AutoFit
   Dim StockTotal As Double
   StockTotal = Cells(2, 7)
   a = 2
   For i = 2 To combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
   'LastRow = Cells(Rows.Count, 1).End(xlUp).Row
       If (Cells(i, 1).Value = Cells(i + 1, 1).Value) Then
           StockTotal = StockTotal + Cells(i + 1, 7).Value
           Else
           Cells(a, 10) = Cells(i, 1)
           Cells(a, 11) = StockTotal
           StockTotal = Cells(i, 7)
           a = a + 1
       End If
   Next i
   combined_sheet.Columns("A:K").AutoFit
   
End Sub

