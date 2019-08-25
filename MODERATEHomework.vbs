Sub moderatehw()

' ---------
' MODERATE
' ---------
' 1. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' 2. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' 3. Conditional formatting that will highlight positive change in green and negative change in red.

' Finds location of <ticker>, <vol>, <open>, <close> and <date> Columns - - -
' Converts Date numbers to Dates - - -
' Holds First Day Opening Pice (FDOP)
' Holds Last Day Closing Price (LDCP)
' Performs % Change Equation on FDOP and LDCP
' Prints % Change in Summary Table
' If % Change is 0+ Then Green
' If % Change is - Then Red

    For Each ws In Worksheets
    
        ' Establishes Last Column
            
        lastcolumn = ws.Range("A1").CurrentRegion.Columns.Count
        
        ' Determines Last Row
            
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        ' Finds location of <ticker>, <vol>, <open>, <close> and <date> columns
        
        Dim datecolumn, tickercolumn, volcolumn, opencolumn, closecolumn As Integer
            
        For j = 1 To lastcolumn
            
            If ws.Cells(1, j).Value = "<ticker>" Then
            tickercolumn = ws.Cells(1, j).Column
                
            Else
                    
                If ws.Cells(1, j).Value = "<vol>" Then
                volcolumn = ws.Cells(1, j).Column
                
                Else
                    
                    If ws.Cells(1, j).Value = "<date>" Then
                    datecolumn = ws.Cells(1, j).Column
                    
                    Else
                    
                        If ws.Cells(1, j).Value = "<open>" Then
                        opencolumn = ws.Cells(1, j).Column
                        
                        Else
                            
                            If ws.Cells(1, j).Value = "<close>" Then
                            closecolumn = ws.Cells(1, j).Column
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
                
            End If
                
        Next j
        
        'Checks if <date> is Already in Date Format
        
        Dim ContainsSlash As Boolean
        ContainsSlash = InStr(1, (ws.Cells(2, datecolumn).Value), "/")
    
        'If <date> is Not in Date Fomat Then...
        If ContainsSlash = False Then
        
        For N = 2 To lastrow
        
                Dim calcyear, calcmonth, calcday, calcdate, finaldate As Date
    
                calcdate = ws.Cells(N, datecolumn).Value
    
                calcyear = Left(calcdate, 4)
                
                calcmonth = Mid(calcdate, 5, 2)
                    
                calcday = Right(calcdate, 2)
        
                finaldate = calcmonth & "/" & calcday & "/" & calcyear
            
                ws.Cells(N, datecolumn).Value = finaldate
        
            Next N
        
        End If
            
        ' Sets Variable for Ticker Name Column Location
        
        Dim TickerNameColumn As Integer
        TickerNameColumn = (lastcolumn + 2)
        
        ' Places "Ticker Name" at the top of TickerNameColumn
        
        ws.Cells(1, TickerNameColumn).Value = "Ticker Name"
            
        ' Sets Variable for YearlyChangeColumn Location
        
        Dim YearlyChangeColumn As Integer
        YearlyChangeColumn = (lastcolumn + 3)
        
        ' Places "Yearly Change" in J1
            
        ws.Cells(1, YearlyChangeColumn).Value = "Yearly Change"
        
        ' Sets Variable for % Change Column Location
        
        Dim PercentChangeColumn As Integer
        PercentChangeColumn = (lastcolumn + 4)
        
        ' Places "Percent Change" at the top of PercentChangeColumn
        
        ws.Cells(1, PercentChangeColumn).Value = "Percent Change"
        
        ' Sets Variable for Ticker Volume Column Location
        
        Dim TickerVolColumn As Integer
        TickerVolColumn = (lastcolumn + 5)
        
        ' Places "Ticker Volume" at the top of TickerVolColumn
        
        ws.Cells(1, TickerVolColumn).Value = "Ticker Volume"
            
        ' Sets Ticker Name As Variable
            
        Dim TickerName As String
            
        ' Sets an Initial Variable for Holding the Total Per Ticker
            
        Dim TickerTotal As Long
        TickerTotal = 0
            
        ' Keeps Track of the Location for Each Ticker Name in the Summary Table
            
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        ' Keeps Track of the YearlyChange for Each Ticker
        
        Dim OpeningSummaryTableRow As Integer
        OpeningTableRow = 3
        
        ' Sets Opening Price as the First Cell in <open> column
        
        Dim Opening As Double
        Opening = ws.Cells(2, opencolumn).Value
        
        ws.Cells(2, YearlyChangeColumn).Value = FirstTickerOpening
       
        ' Loops Through Ticker Volume
        
        For i = 2 To lastrow
        
            ' Checks If We Are Still Within the Same Ticker, If We Are Not...
                    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
            ' Determines the Ticker Name
                    
            TickerName = ws.Cells(i, tickercolumn).Value
                
            ' Sets variables and values for ticker closing
                
            Dim TickerClosing As Double
            TickerClosing = ws.Cells(i, closecolumn).Value
                
            ' Sets Yearly Change Variable and Computation
                
            Dim YearlyChange As Double
                
            YearlyChange = TickerClosing - Opening
                
            ws.Cells(SummaryTableRow, YearlyChangeColumn).Value = YearlyChange
                
            ' Assigns a Color to The Yearly Change Cell, Based on Change
                
            If ws.Cells(SummaryTableRow, YearlyChangeColumn).Value > 0 Then
            ws.Cells(SummaryTableRow, YearlyChangeColumn).Interior.ColorIndex = 4
                
            Else
            
                If ws.Cells(SummaryTableRow, YearlyChangeColumn).Value = 0 Then
                ws.Cells(SummaryTableRow, YearlyChangeColumn).Interior.ColorIndex = 0
                
                Else
                    
                    If ws.Cells(SummaryTableRow, YearlyChangeColumn).Value < 0 Then
                    ws.Cells(SummaryTableRow, YearlyChangeColumn).Interior.ColorIndex = 3
                    
                    End If
                
                End If
                    
            End If
                
            ' Sets Percent Change Variable and Computation and Places Them Value in the Summary Table. Also tells what to do if Opening is zero.
            Dim PercentChange As Double
                
            If Opening = 0 Then
                
            PercentChange = TickerClosing
                    
                Else
                    
                PercentChange = YearlyChange / Opening
                
            End If
                
            ' Formats PercentChange as a Percentage
                
            Dim PercentFormat As String
                
            PercentFormat = FormatPercent(PercentChange, 0)
                
            ws.Cells(SummaryTableRow, PercentChangeColumn).Value = PercentFormat
                
            ' Sets variables and values for ticker opening
                
            Opening = ws.Cells(i + 1, opencolumn).Value
                
            ' Prints ticker closing price in summary table
                
            ws.Cells(SummaryTableRow, PercentChangeColumn).Value = PercentChange
                
            ' Prints ticker opening price in summary table
                
            ws.Cells(SummaryTableRow, YearlyChangeColumn).Value = YearlyChange
                
            ' Adds to the Ticker Total
                    
            TickerTotal = TickerTotal + ws.Cells(i, volcolumn).Value
                
            ' Prints Ticker Name in Summary Table
                    
            ws.Cells(SummaryTableRow, TickerNameColumn).Value = TickerName
    
            ' Prints Ticker Total to the Summary Table
                    
            ws.Cells(SummaryTableRow, TickerVolColumn).Value = TickerTotal
                    
            ' Adds One to the Summary Table Row
                
            SummaryTableRow = SummaryTableRow + 1
                    
            ' Reset the Brand Total
                    
            TickerTotal = 0
                
            Else
                        
                ' Adds to the Ticker Total
                        
                TickerTotal = TickerTotal + ws.Cells(i, TickerVolColumn).Value
                    
            End If
                
        Next i
            
    Next ws

End Sub




