Attribute VB_Name = "Module3"
Sub ExtractRf()

    'On Error Resume Next
    
    'Copy
    Sheets("Adjusted Close Price").Select
    Range("A1", Range("B1").End(xlDown)).Select
    Selection.Copy
        
    'Paste
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Rf"
    Sheets("Rf").Paste
    
    'Replace null to blank
    ActiveSheet.Cells.Replace what:="null", Replacement:="", _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
   'Delete all empty cells row
    Dim Rng As Range
    Set Rng = [B:B].SpecialCells(xlCellTypeBlanks).EntireRow
    Intersect(Rng, Rng).Delete
    
    'Risk-free rate in percentage
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Rf"
    Range("C2").Select

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/100"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C3248")
 
End Sub




