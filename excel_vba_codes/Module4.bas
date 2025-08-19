Attribute VB_Name = "Module4"
Sub RestrictHoldingFormulae()

Dim i As Integer

Sheets("MVF").Activate

For i = 0 To 30
    
    Range("$AL$13").Offset(i, 0).Select
    Selection.FormulaArray = _
        "=MMULT(MMULT(RC[3]:RC[12],Matrix),TRANSPOSE(RC[3]:RC[12]))"
    
    Range("$AN$13").Offset(i, 0).Select
    Selection.FormulaArray = _
        "=MMULT(H.ReturnVec,TRANSPOSE(RC[1]:RC[10]))"
     
    Next i
    
    
End Sub

Sub RestrictHoldings()

'IMPT! Activate Solver - within VBA, "Tools" -> "References" -> "Solver"

Dim i As Integer
'Dim result As Long

Sheets("MVF").Activate

For i = 0 To 30
   
    SolverReset
    SolverOptions Precision:=1E-09
    
    'Sum of weights =1
    SolverAdd CellRef:=Range("$AY$13").Offset(i, 0).Address, Relation:=2, FormulaText:="1"
    
    'Restrict Asset holdings weights
        'Equities
        SolverAdd CellRef:=Range("$AZ$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$AZ$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"
        'Bond
        SolverAdd CellRef:=Range("$BA$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BA$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"
        'Currency
        SolverAdd CellRef:=Range("$BB$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BB$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"
        'Commodities
        SolverAdd CellRef:=Range("$BC$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BC$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"

    'Return to equal to our expectation
    SolverAdd CellRef:=Range("$AK$13").Offset(i, 0).Address, Relation:=2, FormulaText:=Range("$AN$13").Offset(i, 0).Address
    
    'Optimisation problem with minimuim variance and by changing weights
    SolverOk SetCell:=Range("$AL$13").Offset(i, 0).Address, MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$AO$13", "$AX$13").Offset(i, 0).Address, Engine:=1, EngineDesc:="GRG Nonlinear"
    
    'No Short Sales Restriction
    SolverOptions AssumeNonNeg:=False
    SolverSolve (True)

If Range("$AN$13").Offset(i, 0).Value = Range("$AK$13").Offset(i, 0).Value Then

    ElseIf Range("$AY$13").Offset(i, 0).Value = 1 Then
    
    'Equities
    ElseIf Range("$AZ$13").Offset(i, 0).Value >= 0.05 Or Range("$AZ$13").Offset(i, 0).Value <= -0.05 Then
    
    'Bond
    ElseIf Range("$BAZ$13").Offset(i, 0).Value >= 0.05 Or Range("$BA$13").Offset(i, 0).Value <= -0.05 Then
    
    'Currency
    ElseIf Range("$BB$13").Offset(i, 0).Value >= 0.05 Or Range("$BB$13").Offset(i, 0).Value <= -0.05 Then
    
    'Commodities
    ElseIf Range("$BC$13").Offset(i, 0).Value >= 0.05 Or Range("$BC$13").Offset(i, 0).Value <= -0.05 Then
        
    SolverFinish KeepFinal:=1
    
    Else
    
    SolverFinish KeepFinal:=2 'Discard result
    
End If

    Next i
    
End Sub

Sub RestrictShortSalesFormulae()

Dim i As Integer

Sheets("MVF").Activate

For i = 0 To 30
    
    Range("$BF$13").Offset(i, 0).Select
    Selection.FormulaArray = _
        "=MMULT(MMULT(RC[3]:RC[12],Matrix),TRANSPOSE(RC[3]:RC[12]))"
    
    Range("$BH$13").Offset(i, 0).Select
    Selection.FormulaArray = _
        "=MMULT(H.ReturnVec,TRANSPOSE(RC[1]:RC[10]))"
     
    Next i
    
    
End Sub

Sub RestrictShortSales()

'IMPT! Activate Solver - within VBA, "Tools" -> "References" -> "Solver"

Dim i As Integer
'Dim result As Long

Sheets("MVF").Activate

For i = 5 To 30
   
    SolverReset
    SolverOptions Precision:=1E-09
    
    'Sum of weights =1
    SolverAdd CellRef:=Range("$BS$13").Offset(i, 0).Address, Relation:=2, FormulaText:="1"
    
    'Restrict Asset holdings weights
        'Equities
        SolverAdd CellRef:=Range("$BT$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BT$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"

        'Bond
        SolverAdd CellRef:=Range("$BU$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BU$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"
        'Currency
        SolverAdd CellRef:=Range("$BV$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BV$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"
        'Commodities
        SolverAdd CellRef:=Range("$BW$13").Offset(i, 0).Address, Relation:=3, FormulaText:="$BC$9"
        SolverAdd CellRef:=Range("$BW$13").Offset(i, 0).Address, Relation:=1, FormulaText:="$BB$9"

    'Return to equal to our expectation
    SolverAdd CellRef:=Range("$BE$13").Offset(i, 0).Address, Relation:=2, FormulaText:=Range("$BH$13").Offset(i, 0).Address
    
    'Optimisation problem with minimuim variance and by changing weights
    SolverOk SetCell:=Range("$BF$13").Offset(i, 0).Address, MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$BI$13", "$BR$13").Offset(i, 0).Address, Engine:=1, EngineDesc:="GRG Nonlinear"
    
    'Short Sales Restricted
    SolverOptions AssumeNonNeg:=True
    SolverSolve (True)

If Range("$BH$13").Offset(i, 0).Value = Range("$BE$13").Offset(i, 0).Value Then

    ElseIf Range("$BS$13").Offset(i, 0).Value = 1 Then
    
    'Equities
    ElseIf Range("$BT$13").Offset(i, 0).Value >= 0.05 Or Range("$BT$13").Offset(i, 0).Value <= -0.05 Then
    
    'Bond
    ElseIf Range("$BU$13").Offset(i, 0).Value >= 0.05 Or Range("$BU$13").Offset(i, 0).Value <= -0.05 Then
    
    'Currency
    ElseIf Range("$BV$13").Offset(i, 0).Value >= 0.05 Or Range("$BV$13").Offset(i, 0).Value <= -0.05 Then
    
    'Commodities
    ElseIf Range("$BW$13").Offset(i, 0).Value >= 0.05 Or Range("$BW$13").Offset(i, 0).Value <= -0.05 Then
        
    SolverFinish KeepFinal:=1
    
    Else
    
    SolverFinish KeepFinal:=2 'Discard result
    
End If

    Next i
    
End Sub

