VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    With Range("A" & firstTickerRow & ":A" & Rows.Count).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With Range("A" & firstTickerRow - 1 & ":A" & Rows.Count)
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    lastRowColA = Cells(Rows.Count, "a").End(xlUp).Row

    With Range("A" & firstTickerRow - 1 & ":A" & lastRowColA)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    With Range("A" & firstTickerRow & ":A" & lastRowColA).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

End Sub

