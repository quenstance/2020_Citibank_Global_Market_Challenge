Attribute VB_Name = "Module2"
Option Explicit
Sub CleanData()

    'On Error Resume Next
    
    'Copy
    Sheets("Adjusted Close Price").Select
    Sheets("Adjusted Close Price").UsedRange.Copy
        
    'Paste Values only
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Cleaned Data"
    Sheets("Cleaned Data").Paste
    
    'Copy and PasteSpecial a Range
    Range("N2", Range("Q2").End(xlDown)).Copy
    Range("N2").PasteSpecial Paste:=xlPasteValues

    'Replace null to blank
    ActiveSheet.Cells.Replace what:="null", Replacement:="", _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    
    'Unselect
    Application.CutCopyMode = False
    Range("A1").Select
    
    'Delete all empty cells row
    
    Dim Rng As Range
    Set Rng = Range("A1", Range("A1").End(xlDown).End(xlToRight)).SpecialCells(xlCellTypeBlanks).EntireRow
    Intersect(Rng, Rng).Delete
    
    'Number format
    Range("B2", Range("B2").End(xlDown).End(xlToRight)).Select
    Selection.NumberFormat = "0.00"
    


End Sub

