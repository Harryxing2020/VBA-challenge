VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub delete_cell()

Dim CurrentSheet As Worksheet

' clear the colume for debug the code
For Each CurrentSheet In Worksheets
    CurrentSheet.Columns("I:Q").Delete
    

Next CurrentSheet


End Sub

Sub test_Excelhomework2()

Dim CurrentSheet As Worksheet

Dim currentRow As Long
Dim count As Long
Dim currentTicker As String
Dim countStock As Long
Dim yearChange, yearBegin, yearClose As Double
Dim percentChang As Double
Dim total_stock_volume As Double
Dim greatIncrease, greatestDecrease, GreatestVolume As Double
Dim GreatIncreaseTicker, greatestDecreaseTicker, GreatestVolumeTicker As String


'setup loop for each worksheet
For Each CurrentSheet In Worksheets
    
    CurrentSheet.Select
    CurrentSheet.Activate
    'Select and activate the current sheet
    
    'Obtain the name of worksheet
    currentTicker = CurrentSheet.Name
    
    'inital the value of currentTicker
    currentTicker = ""
    countStock = 2
    
    'inital the value of variable
    
    yearChange = yearBegin = yearClose = percentChang = total_stock_volume = 0
    greatIncrease = greatestDecrease = GreatestVolume = 0
    
    
    
    
    'Get the total number of row
    currentRow = CurrentSheet.Cells(Rows.count, 1).End(xlUp).Row
        
    
    'put the table header
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "YearChange"
    Cells(1, "k").Value = "percentageChang"
    Cells(1, "L").Value = "Total stock volume"
    
    Cells(2, "O").Value = "Greatest % increase"
    Cells(3, "O").Value = "Greatest % decrease"
    Cells(4, "O").Value = "Greatest total volume"
    
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    'adjust the width of each column
    Columns("I").ColumnWidth = Len("Ticker")
    Columns("J").ColumnWidth = Len("YearChange")
    Columns("K").ColumnWidth = Len("percentageChang")
    Columns("L").ColumnWidth = Len("Total stock volume")
    
    Columns("O").ColumnWidth = Len("Greatest total volume")
    
    Columns("P").ColumnWidth = Len("Ticker")
    Columns("Q").ColumnWidth = 10
    
    
    
    ' Get each row in the current worksheet
    For count = 2 To currentRow
        
        ' add the stock volume for each stock
        total_stock_volume = total_stock_volume + Cells(count, "G")
        
        If currentTicker = "" Then
            ' get the stock ticker
            currentTicker = Cells(count, "A")
            ' put the current ticker to output cell
            Cells(countStock, "I").Value = currentTicker
            ' get begin year
            yearBegin = Cells(count, "C")
            
            ' if the ticker name of next row is different from current ticker, then need to calculate the current stock's
            'yearchange and percentage
        ElseIf currentTicker <> Cells(count + 1, "A") Then
            
            yearClose = Cells(count, "F")
            yearChange = yearClose - yearBegin
            
            Cells(countStock, "J").Value = yearChange
            
            If yearBegin <> 0 Then
                percentChang = yearChange / yearBegin
            
            Else
                percentChang = 0
            End If
            
            
            ' assign the percentchange to current cell.
            Cells(countStock, "K").Value = percentChang
            
            ' formating the cell
            Cells(countStock, "K").NumberFormat = "0.00%"
            
            'Interior.ColorIndex =4
            
            ' formating the backgroud color of the cell
            If percentChang > 0 Then
                Cells(countStock, "K").Interior.ColorIndex = 4
            ElseIf percentChang < 0 Then
                Cells(countStock, "K").Interior.ColorIndex = 3
            End If
            
            ' assign the total stock volume to current cell.
            Cells(countStock, "L").Value = total_stock_volume
            
            ' if the increase percentage is greater than our your record, then we record the current stock
            If percentChang > greatIncrease Then
            
                GreatIncreaseTicker = Cells(count, "A")
                greatIncrease = percentChang

            End If
            
            ' if the decrease percentage is greater than our your record, then we record the current stock
            If percentChang < greatestDecrease Then
            
                greatestDecreaseTicker = Cells(count, "A")
                greatestDecrease = percentChang

            End If
            
            ' if the total stock vo;ume is greater than our your record, then we record the current stock
            
            If GreatestVolume < total_stock_volume Then
            
                GreatestVolumeTicker = Cells(count, "A")
                GreatestVolume = total_stock_volume

            End If
            
            currentTicker = ""
            countStock = countStock + 1
            total_stock_volume = 0
        End If
        
    Next count
    
    ' assign the value of data
    Cells(2, "P").Value = GreatIncreaseTicker
    Cells(2, "Q").Value = greatIncrease
    Cells(3, "P").Value = greatestDecreaseTicker
    Cells(3, "Q").Value = greatestDecrease
    
    Cells(4, "P").Value = GreatestVolumeTicker
    Cells(4, "Q").Value = GreatestVolume
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    
Next CurrentSheet


End Sub




