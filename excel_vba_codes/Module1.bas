Attribute VB_Name = "Module1"
Option Explicit
Public Const firstTickerRow As Integer = 13

Sub DownloadData()

    Dim frequency As String
    Dim lastRow As Integer
    Dim lastErrorRow As Integer
    Dim lastSuccessRow As Integer
    Dim stockTicker As String
    Dim numStockErrors As Integer
    Dim numStockSuccess As Integer
    Dim startDate As String
    Dim endDate As String
    Dim ticker As Integer
    
    Dim crumb As String
    Dim cookie As String
    Dim validCookieCrumb As Boolean
    
    Dim sortOrderComboBox As Shape
 
    numStockErrors = 0
    numStockSuccess = 0

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    lastErrorRow = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
    lastSuccessRow = ActiveSheet.Cells(Rows.Count, "E").End(xlUp).Row

    ClearErrorList lastErrorRow
    ClearSuccessList lastSuccessRow

    lastRow = ActiveSheet.Cells(Rows.Count, "a").End(xlUp).Row
    frequency = Worksheets("GetData").Range("b7")
    
    'Convert user-specified calendar dates to Unix time
    '***************************************************
    startDate = (Sheets("GetData").Range("startDate") - DateValue("January 1, 1970")) * 86400
    endDate = (Sheets("GetData").Range("endDate") - DateValue("January 1, 1970")) * 86400
    '***************************************************
    
    'Set date retrieval frequency
    '***************************************************
    If Worksheets("GetData").Range("frequency") = "d" Then
        frequency = "1d"
    ElseIf Worksheets("GetData").Range("frequency") = "w" Then
        frequency = "1wk"
    ElseIf Worksheets("GetData").Range("frequency") = "m" Then
        frequency = "1mo"
    End If
    '***************************************************

    'Delete all sheets apart from GetData sheet
    '***************************************************
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name <> "GetData" And ws.Name <> "FundX" Then ws.Delete
    Next
    '***************************************************

    'Get cookie and crumb
    '***************************************************
    Call getCookieCrumb(crumb, cookie, validCookieCrumb)
    If validCookieCrumb = False Then
        GoTo ErrorHandler:
    End If
    '***************************************************

    'Loop through all tickers
    For ticker = firstTickerRow To lastRow

        stockTicker = Worksheets("GetData").Range("$a$" & ticker)

        If stockTicker = "" Then
            GoTo NextIteration
        End If

        'Create a sheet for each ticker
        '***************************************************
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = stockTicker
        Cells(1, 1) = "Stock Quotes for " & stockTicker
        '***************************************************

        'Get financial data from Yahoo and write into each sheet
        'getCookieCrumb() must be run before running getYahooFinanceData()
        '***************************************************
        Call getYahooFinanceData(stockTicker, startDate, endDate, frequency, cookie, crumb)
        '***************************************************
        
        
        'Populate success or fail lists
        '***************************************************
        lastRow = Sheets(stockTicker).UsedRange.Row - 2 + Sheets(stockTicker).UsedRange.Rows.Count

        If lastRow < 3 Then
            Sheets(stockTicker).Delete
            numStockErrors = numStockErrors + 1
            ErrorList stockTicker, numStockErrors
            GoTo NextIteration
        Else
            numStockSuccess = numStockSuccess + 1
            If Left(stockTicker, 1) = "^" Then
                SuccessList Replace(stockTicker, "^", ""), numStockSuccess
            Else
                SuccessList stockTicker, numStockSuccess
            End If
        End If
        '***************************************************

        'Set the preferred date format
        '***************************************************
        Range("a2:a" & lastRow).NumberFormat = "yyyy-mm-dd;@"
        '***************************************************
        
        'Sort by oldest date first or newest date first
        '***************************************************
        Set sortOrderComboBox = Sheets("GetData").Shapes("SortOrderDropDown")
        With sortOrderComboBox.ControlFormat
            If .List(.Value) = "Oldest First" Then
                Call SortByDate(stockTicker, "oldest")
            ElseIf .List(.Value) = "Newest First" Then
                Call SortByDate(stockTicker, "newest")
            End If
        End With
        '***************************************************
        
        'Clean up sheet names
        '***************************************************
        'Remove initial ^ in ticker names from Sheets
        'If Left(stockTicker, 1) = "^" Then
            'ActiveSheet.Name = Replace(stockTicker, "^", "")
        'Else
            'ActiveSheet.Name = stockTicker
        'End If

        'Remove hyphens in ticker names from Sheet names, otherwise error in collation
        'If InStr(stockTicker, "-") > 0 Then
            'ActiveSheet.Name = Replace(stockTicker, "-", "")
        'End If
        '***************************************************

NextIteration:
    Next ticker
    
    'Process export and collation
    '***************************************************
    If Sheets("GetData").Shapes("WriteToCSVCheckBox").ControlFormat.Value = xlOn Then
        On Error GoTo ErrorHandler:
        Call CopyToCSV
    End If

    If Sheets("GetData").Shapes("CollateDataCheckBox").ControlFormat.Value = xlOn Then
        On Error GoTo ErrorHandler:
        Call CollateData
    End If
    '***************************************************
ErrorHandler:

    Worksheets("GetData").Select
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub SortByDate(ticker As String, order As String)
    
    Dim firstRow As Integer
    Dim lastRow As Integer
    Dim sortType As Variant
    
    lastRow = Sheets(ticker).UsedRange.Rows.Count
    firstRow = 2
    
    If order = "oldest" Then
       sortType = xlAscending
    Else
       sortType = xlDescending
    End If
    
    Worksheets(ticker).Sort.SortFields.Clear
    Worksheets(ticker).Sort.SortFields.Add Key:=Sheets(ticker).Range("A" & firstRow & ":A" & lastRow), _
        SortOn:=xlSortOnValues, order:=sortType, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ticker).Sort
        .SetRange Range("A" & firstRow & ":G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Sub CollateData()

    Dim ws As Worksheet
    Dim i As Integer
    Dim maxRow As Integer
    Dim maxTickerWS As Worksheet

    maxRow = 0
    For Each ws In Worksheets
        If ws.Name <> "GetData" Then
            If ws.UsedRange.Rows.Count > maxRow Then
                maxRow = ws.UsedRange.Rows.Count
                Set maxTickerWS = ws
            End If
        End If
    Next

    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Adjusted Close Price"


'****************************************
    i = 1
  
    maxTickerWS.Range("A2", "a" & maxRow).Copy Destination:=Sheets("Adjusted Close Price").Cells(1, i)
    maxTickerWS.Range("f2", "f" & maxRow).Copy Destination:=Sheets("Adjusted Close Price").Cells(1, i + 1)
    Sheets("Adjusted Close Price").Cells(1, i + 1) = maxTickerWS.Name
    

    i = i + 2

    For Each ws In Worksheets

        If ws.Name <> "GetData" And ws.Name <> "FundX" And ws.Name <> maxTickerWS.Name And ws.Name <> "Adjusted Close Price" Then

            Sheets("Adjusted Close Price").Cells(1, i) = ws.Name
            Sheets("Adjusted Close Price").Range(Sheets("Adjusted Close Price").Cells(2, i), Sheets("Adjusted Close Price").Cells(maxRow - 1, i)).Formula = _
                "=vlookup(A2," & ws.Name & "!A$2:G$" & maxRow & ",6,0)"



'****************************************
            i = i + 1

        End If
    Next

    On Error Resume Next


    Sheets("Adjusted Close Price").UsedRange.SpecialCells(xlFormulas, xlErrors).Clear


    Sheets("Adjusted Close Price").UsedRange.Value = Sheets("Adjusted Close Price").UsedRange.Value
    On Error GoTo 0


    Sheets("Adjusted Close Price").Columns("A:A").EntireColumn.AutoFit
End Sub

Sub SuccessList(ByVal stockTicker As String, ByVal numStockSuccess As Integer)

    Sheets("GetData").Range("E" & 12 + numStockSuccess) = stockTicker

    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeLeft).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeTop).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeBottom).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeRight).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlDiagonalUp).LineStyle = xlNone

    With Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Borders(xlInsideHorizontal).LineStyle = xlNone

    With Sheets("GetData").Range("E12:E" & 12 + numStockSuccess).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

End Sub

Sub ErrorList(ByVal stockTicker As String, ByVal numStockErrors As Integer)

    Sheets("GetData").Range("C" & 12 + numStockErrors) = stockTicker

    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlDiagonalUp).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeLeft).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeTop).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeBottom).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeRight).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlDiagonalDown).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlDiagonalUp).LineStyle = xlNone

    With Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlInsideVertical).LineStyle = xlNone
    Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Borders(xlInsideHorizontal).LineStyle = xlNone

    With Sheets("GetData").Range("C12:C" & 12 + numStockErrors).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

End Sub

Sub ClearErrorList(ByVal lastErrorRow As Integer)
    If lastErrorRow > 12 Then
        Worksheets("GetData").Range("C13:C" & lastErrorRow).Clear
        With Sheets("GetData").Range("C12").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("C12").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("C12").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("C12").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End If
End Sub

Sub ClearSuccessList(ByVal lastSuccessRow As Integer)
    If lastSuccessRow > 12 Then
        Worksheets("GetData").Range("E13:F" & lastSuccessRow).Clear
        With Sheets("GetData").Range("E12").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("E12").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("E12").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Sheets("GetData").Range("E12").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End If
End Sub

Sub CopyToCSV()

    Dim MyPath As String
    Dim MyFileName As String
    Dim dateFrom As Date
    Dim dateTo As Date
    Dim frequency As String
    Dim ws As Worksheet
    Dim ticker As String

    dateFrom = Worksheets("GetData").Range("$b$5")
    dateTo = Worksheets("GetData").Range("$b$6")
    frequency = Worksheets("GetData").Range("$b$7")
    MyPath = Worksheets("GetData").Range("$b$8")

    For Each ws In Worksheets
        If ws.Name <> "GetData" Then
            ticker = ws.Name
            MyFileName = ticker & " " & Format(dateFrom, "dd-mm-yyyy") & " - " & Format(dateTo, "dd-mm-yyyy") & " " & frequency
            If Not Right(MyPath, 1) = "\" Then MyPath = MyPath & "\"
            If Not Right(MyFileName, 4) = ".csv" Then MyFileName = MyFileName & ".csv"
            Sheets(ticker).Copy
            With ActiveWorkbook
                .SaveAs Filename:= _
                    MyPath & MyFileName, _
                    FileFormat:=xlCSV, _
                    CreateBackup:=False
                .Close False
            End With
        End If
    Next
End Sub

Sub getCookieCrumb(crumb As String, cookie As String, validCookieCrumb As Boolean)

    Dim i As Integer
    Dim str As String
    Dim crumbStartPos As Long
    Dim crumbEndPos As Long
    Dim objRequest
 
    validCookieCrumb = False
    
    For i = 0 To 5  'ask for a valid crumb 5 times
        Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        With objRequest
            .Open "GET", "https://finance.yahoo.com/lookup?s=bananas", False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
            .send
            .waitForResponse (10)
            cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
            crumbStartPos = InStrRev(.ResponseText, """crumb"":""") + 9
            crumbEndPos = crumbStartPos + 11 'InStr(crumbStartPos, .ResponseText, """", vbBinaryCompare)
            crumb = Mid(.ResponseText, crumbStartPos, crumbEndPos - crumbStartPos)

        End With
        
        If Len(crumb) = 11 Then 'a valid crumb is 11 characters long
            validCookieCrumb = True
            Exit For
        End If:
        

    Next i
    
End Sub

Sub getYahooFinanceData(stockTicker As String, startDate As String, endDate As String, frequency As String, cookie As String, crumb As String)
    Dim resultFromYahoo As String
    Dim objRequest
    Dim csv_rows() As String
    Dim resultArray As Variant
    Dim nColumns As Integer
    Dim iRows As Integer
    Dim CSV_Fields As Variant
    Dim iCols As Integer
    Dim tickerURL As String

    'Construct URL
    '***************************************************
    tickerURL = "https://query1.finance.yahoo.com/v7/finance/download/" & stockTicker & _
        "?period1=" & startDate & _
        "&period2=" & endDate & _
        "&interval=" & frequency & "&events=history" & "&crumb=" & crumb
    'Sheets("GetData").Range("K" & ticker - 1) = tickerURL
    '***************************************************
               
    'Get data from Yahoo
    '***************************************************
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .setRequestHeader "Cookie", cookie
        .send
        .waitForResponse
        resultFromYahoo = .ResponseText
    End With
    '***************************************************
        
    'Parse returned string into an array
    '***************************************************
    nColumns = 6 'number of columns minus 1  (date, open, high, low, close, adj close, volume)
    csv_rows() = Split(resultFromYahoo, Chr(10))
    ReDim resultArray(0 To UBound(csv_rows), 0 To nColumns) As Variant
     
    For iRows = LBound(csv_rows) To UBound(csv_rows)
        CSV_Fields = Split(csv_rows(iRows), ",")
        If UBound(CSV_Fields) > nColumns Then
            nColumns = UBound(CSV_Fields)
            ReDim Preserve resultArray(0 To UBound(csv_rows), 0 To nColumns) As Variant
        End If
    
        For iCols = LBound(CSV_Fields) To UBound(CSV_Fields)
            If IsNumeric(CSV_Fields(iCols)) Then
                resultArray(iRows, iCols) = Val(CSV_Fields(iCols))
            ElseIf IsDate(CSV_Fields(iCols)) Then
                resultArray(iRows, iCols) = CDate(CSV_Fields(iCols))
            Else
                resultArray(iRows, iCols) = CStr(CSV_Fields(iCols))
            End If
        Next
    Next
 
    'Write results into worksheet for ticker
    Sheets(stockTicker).Range("A2").Resize(UBound(resultArray, 1) + 1, UBound(resultArray, 2) + 1).Value = resultArray
    '***************************************************
    
End Sub



