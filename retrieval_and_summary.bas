Attribute VB_Name = "Module111"
Sub retrieval_and_summary()

    For Each ws In Worksheets
        ws.Select
        Call retrieval
        Call summary
    Next
    
End Sub
Sub retrieval()
    
    'sort table
    Dim n_row As Long
    n_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    range_n_row = WorksheetFunction.Text(n_row, "#")
    
    ActiveSheet.Sort.SortFields.Clear
    
    Range("A1:G" & range_n_row).Sort Key1:=Range("A1"), header:=xlYes, Key2:=Range("B1"), header:=xlYes
    
    ' PART 1 - retrieval table
    
    ' set variables used in count
    Dim sum_t_vol As LongLong
    Dim initial_row, n_ticker, open_price, close_price As Double

    ' report header
    Dim header(0 To 3) As String
        header(0) = "Ticker"
        header(1) = "Yearly Change"
        header(2) = "Percent Change"
        header(3) = "Total Stock Volume"
    Range(Cells(1, 9), Cells(1, 12)) = header
    
    Dim current_ticker As String
           
    Dim i As Long
        
    For i = 1 To n_row
        
        ' set initial value
        If i = 1 Then
            sum_t_vol = 0
            close_price = Empty
            
            initial_row = 1
            n_ticker = 1
            current_ticker = Cells(i + 1, 1).Value
            open_price = Cells(i + 1, 3).Value
        Else
            
            ' add value for current row
            sum_t_vol = sum_t_vol + Cells(i, 7).Value
            
            ' when next line is a different tricker, add current data to retrieval tabel and reset for new ticker
            If Cells(i + 1, 1).Value <> current_ticker Then
            
                ' Current Ticker Name
                Cells(initial_row + n_ticker, 9) = current_ticker
                
                ' Yearly change = close at last day - open at first day
                close_price = Cells(i, 6).Value
                Cells(initial_row + n_ticker, 10) = close_price - open_price
                
                ' add format to yearly change
                If close_price - open_price < 0 Then
                    Cells(initial_row + n_ticker, 10).Interior.ColorIndex = "3"
                ElseIf close_price - open_price > 0 Then
                    Cells(initial_row + n_ticker, 10).Interior.ColorIndex = "4"
                End If
                                
                ' % change
                Cells(initial_row + n_ticker, 11) = FormatPercent((close_price - open_price) / open_price)
                
                ' Total Stock Volume, no scientific notation
                Cells(initial_row + n_ticker, 12) = sum_t_vol
                Cells(initial_row + n_ticker, 12).NumberFormat = "0"
                
                ' name for new ticker
                current_ticker = Cells(i + 1, 1).Value
                n_ticker = n_ticker + 1
                
                ' reset values
                sum_t_vol = 0
                open_price = Cells(i + 1, 3).Value
                close_price = Empty
                
            End If
        
        End If
        
    Next i
    
End Sub

Sub summary()
            
    ' PART 2 - summarize retrieval table
    
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    
    Dim n_row, row_initial, max_per, min_per As Double
        n_row = Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim max_vol As LongLong
    
    Dim ticker_max_per, ticker_min_per, ticker_max_vol As String
    
    Dim i As Integer
    
    For i = 2 To n_row - 1 'stop at n_row-1 the next row is blank
        
        ' use row 2 as initial values, compare with next row, if the value is larger than max or less than min then replace
        If i = 2 Then
            
            ' set initial ticker name
            ticker_max_per = Cells(i, 9).Value
            ticker_min_per = Cells(i, 9).Value
            ticker_max_vol = Cells(i, 9).Value
            
            ' set initial value
            max_per = Cells(i, 11).Value
            min_per = Cells(i, 11).Value
            max_vol = Cells(i, 12).Value
        
        Else
            
            ' compare next %, replace max or min. question, what if same?
            If Cells(i + 1, 11).Value > max_per Then
                max_per = Cells(i + 1, 11).Value
                ticker_max_per = Cells(i + 1, 9).Value
                
            ElseIf Cells(i + 1, 11).Value < min_per Then
                min_per = Cells(i + 1, 11).Value
                ticker_min_per = Cells(i + 1, 9).Value
                
            End If
            
            ' compare next vol, replace if larger than max
            If Cells(i + 1, 12).Value > max_vol Then
                max_vol = Cells(i + 1, 12).Value
                ticker_max_vol = Cells(i + 1, 9).Value
                
            End If
        
        End If
        
    Next i
    
    ' write result to table
    Cells(2, 16).Value = ticker_max_per
    Cells(2, 17).Value = FormatPercent(max_per)
    
    Cells(3, 16).Value = ticker_min_per
    Cells(3, 17).Value = FormatPercent(min_per)
    
    Cells(4, 16).Value = ticker_max_vol
    Cells(4, 17).Value = max_vol
    Cells(4, 17).NumberFormat = "0"
    
End Sub





