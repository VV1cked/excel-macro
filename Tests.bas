Attribute VB_Name = "Tests"
Option Explicit

' ===========================
' Tests.bas (ExternSystems + new tp/Wi rules)
' ===========================

'===========================
' Basic smoke tests (kept)
'===========================

Sub Test_FailureCalc()
    Dim result As Double
    result = CalcFailure("SYS6", "ALL")
    Debug.Print "CalcFailure(SYS6, ALL) = " & result
End Sub

Sub TestCalcFailure_SYS6()
    Dim result As Double
    Dim FuncName As String
    Dim stage As Variant
    
    FuncName = "SYS6"
    stage = 0
    
    On Error GoTo ErrHandler
    
    Debug.Print "=== Start CalcFailure test: " & FuncName & ", Stage=" & stage & " ==="
    result = CalcFailure(FuncName, stage)
    Debug.Print "Result = " & result
    Debug.Print "=== End test ==="
    Exit Sub
    
ErrHandler:
    Debug.Print "Error: " & Err.Number & " - " & Err.Description
    MsgBox "Error in CalcFailure test: " & Err.Description, vbCritical
End Sub

'===========================
' Helpers for tests
'===========================

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Function FirstFunctionName() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Functions")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim nm As String
        nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nm) > 0 Then
            FirstFunctionName = nm
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 900, , "Не удалось найти ни одной функции на листе Functions"
End Function

Private Function FirstTpCell() As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Elements")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 3).Value) Then
            Set FirstTpCell = ws.Cells(r, 3)
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 901, , "Не удалось найти числовую ячейку tp в Elements!C"
End Function

Private Sub SetTpAndReinit(ByVal tp As Double)
    Dim c As Range
    Set c = FirstTpCell()
    c.Value = tp

    InitGlobals
    m_CallStack.RemoveAll
End Sub

Private Sub AssertTrue(ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then Err.Raise vbObjectError + 910, , "ASSERT TRUE FAILED: " & msg
End Sub

Private Sub AssertNear(ByVal a As Double, ByVal b As Double, ByVal relTol As Double, ByVal msg As String)
    If a = 0 And b = 0 Then Exit Sub

    Dim denom As Double
    denom = IIf(Abs(b) > Abs(a), Abs(b), Abs(a))
    If denom = 0 Then Err.Raise vbObjectError + 911, , "ASSERT NEAR FAILED (both near 0?): " & msg

    If Abs(a - b) / denom > relTol Then
        Err.Raise vbObjectError + 912, , msg & " (a=" & a & ", b=" & b & ", rel=" & Abs(a - b) / denom & ")"
    End If
End Sub

Private Function IsExternID(ByVal id As Long) As Boolean
    On Error Resume Next
    If m_ExternByID Is Nothing Then
        IsExternID = False
    Else
        IsExternID = m_ExternByID.Exists(id)
    End If
    On Error GoTo 0
End Function

Private Function TermLambdaCount(ByVal t As CTerm) As Long
    Dim idsV As Variant
    idsV = t.FactorIDs
    If IsEmpty(idsV) Then Exit Function

    Dim c As Long, i As Long
    c = 0
    For i = LBound(idsV) To UBound(idsV)
        If Not IsExternID(CLng(idsV(i))) Then c = c + 1
    Next i
    TermLambdaCount = c
End Function

Private Function ExprHasAnyLambda(ByVal e As CExpr) As Boolean
    Dim terms() As CTerm
    terms = e.GetTerms()
    If (Not Not terms) = 0 Then Exit Function

    Dim i As Long
    For i = LBound(terms) To UBound(terms)
        If TermLambdaCount(terms(i)) > 0 Then
            ExprHasAnyLambda = True
            Exit Function
        End If
    Next i
End Function

Private Function ExprAllTermsSameLambdaCount(ByVal e As CExpr, ByRef outCount As Long) As Boolean
    Dim terms() As CTerm
    terms = e.GetTerms()
    If (Not Not terms) = 0 Then Exit Function

    Dim i As Long
    outCount = TermLambdaCount(terms(LBound(terms)))

    For i = LBound(terms) To UBound(terms)
        If TermLambdaCount(terms(i)) <> outCount Then
            ExprAllTermsSameLambdaCount = False
            Exit Function
        End If
    Next i

    ExprAllTermsSameLambdaCount = True
End Function

' ---- NEW: detects tp token in symbolic latex under multiple possible spellings ----
Private Function ContainsTpSymbolic(ByVal s As String) As Boolean
    Dim norm As String
    norm = Replace(s, " ", "")
    
    ' Accept:
    '   t_p  (latin p)
    '   t_п  (cyrillic pe)
    ' plus a couple of fallback variants in case templates omit underscore
    ContainsTpSymbolic = _
        (InStr(1, norm, "t_p", vbTextCompare) > 0) Or _
        (InStr(1, norm, "t_п", vbTextCompare) > 0) Or _
        (InStr(1, norm, "tp", vbTextCompare) > 0) Or _
        (InStr(1, norm, "tп", vbTextCompare) > 0)
End Function

' Find a function whose expression is exactly a single extern name (token only)
Private Function FindSingleExternFunction(ByRef outFuncName As String, ByRef outExternName As String) As Boolean
    If Not SheetExists("Functions") Then Exit Function
    If Not SheetExists("ExternSystems") Then Exit Function

    InitGlobals
    If m_ExternByID Is Nothing Then Exit Function
    If m_ExternByID.Count = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Functions")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim fName As String
        fName = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(fName) = 0 Then GoTo NextRow

        Dim expr As String
        expr = Trim$(CStr(ws.Cells(r, 2).Value))
        expr = Replace(expr, " ", "")
        If Len(expr) = 0 Then GoTo NextRow

        If Left$(expr, 1) = "(" And Right$(expr, 1) = ")" Then
            expr = Mid$(expr, 2, Len(expr) - 2)
        End If

        If InStr(1, expr, "+", vbTextCompare) > 0 Then GoTo NextRow
        If InStr(1, expr, "*", vbTextCompare) > 0 Then GoTo NextRow

        Dim id As Long
        id = GetID(expr)
        If IsExternID(id) Then
            outFuncName = fName
            outExternName = expr
            FindSingleExternFunction = True
            Exit Function
        End If

NextRow:
    Next r
End Function


'===========================
' Updated Tests
'===========================

Public Sub Test_Tp_Affects_CalcFailure_By_LambdaCount()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    InitGlobals
    Dim e As CExpr
    Set e = EvalFunction(fName)
    If e Is Nothing Then Err.Raise vbObjectError + 920, , "EvalFunction returned Nothing for '" & fName & "'"

    Dim nLam As Long
    If Not ExprAllTermsSameLambdaCount(e, nLam) Then
        Debug.Print "SKIP: Test_Tp_Affects_CalcFailure_By_LambdaCount (function '" & fName & "' has mixed lambda-count terms)"
        Exit Sub
    End If

    Dim stage As Variant
    stage = 0

    Dim tp1 As Double, tp2 As Double
    tp1 = 0.5
    tp2 = 1#

    SetTpAndReinit tp1
    Dim v1 As Double
    v1 = CalcFailure(fName, stage)

    SetTpAndReinit tp2
    Dim v2 As Double
    v2 = CalcFailure(fName, stage)

    If nLam = 0 Then
        AssertNear v1, v2, 0.000001, "CalcFailure must not depend on tp when there are no lambdas (f=" & fName & ")"
        Debug.Print "OK: CalcFailure does not depend on tp (no lambdas) for f=" & fName
        Exit Sub
    End If

    Dim expected As Double
    expected = (tp1 / tp2) ^ nLam

    AssertNear v1, v2 * expected, 0.000001, "CalcFailure does not scale as tp^(#lambda) (f=" & fName & ", nLambda=" & nLam & ")"
    Debug.Print "OK: CalcFailure scales by tp^(#lambda) for f=" & fName & ", nLambda=" & nLam
    Exit Sub

EH:
    Debug.Print "FAIL: Test_Tp_Affects_CalcFailure_By_LambdaCount: " & Err.Description
    MsgBox "FAIL: Test_Tp_Affects_CalcFailure_By_LambdaCount" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub Test_RewriteFailure_Includes_tp_only_when_lambdas_exist()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    SetTpAndReinit 0.5

    Dim e As CExpr
    Set e = EvalFunction(fName)

    Dim hasLam As Boolean
    hasLam = ExprHasAnyLambda(e)

    Dim s As String
    s = RewriteFailure(fName, 0)

    AssertTrue Left$(s, 3) = "Q_{", "RewriteFailure must start with 'Q_{'"

    If hasLam Then
        AssertTrue ContainsTpSymbolic(s), "RewriteFailure must contain t_p or t_п when lambdas exist"
    Else
        AssertTrue Not ContainsTpSymbolic(s), "RewriteFailure must NOT contain tp when no lambdas exist"
    End If

    Debug.Print "OK: RewriteFailure tp presence matches lambda presence for f=" & fName
    Exit Sub

EH:
    Debug.Print "FAIL: Test_RewriteFailure_Includes_tp_only_when_lambdas_exist: " & Err.Description
    MsgBox "FAIL: Test_RewriteFailure_Includes_tp_only_when_lambdas_exist" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub Test_SubstituteFailure_Includes_tp_numeric_only_when_lambdas_exist()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    Dim tp As Double
    tp = 0.5
    SetTpAndReinit tp

    Dim e As CExpr
    Set e = EvalFunction(fName)

    Dim hasLam As Boolean
    hasLam = ExprHasAnyLambda(e)

    Dim s As String
    s = SubstituteFailure(fName, 0)

    Debug.Print "SubstituteFailure output:"
    Debug.Print s

    AssertTrue Left$(s, 3) = "Q_{", "SubstituteFailure must start with 'Q_{'"

    Dim norm As String
    norm = Replace(s, " ", "")

    Dim tpPlain As String
    tpPlain = Format$(tp, "0.############")
    tpPlain = Replace(tpPlain, " ", "")

    Dim av As Double, exp As Long, mant As Double
    av = Abs(tp)
    exp = Fix(Log(av) / Log(10#))
    mant = tp / (10# ^ exp)

    Dim mantStr As String
    mantStr = Format$(mant, "0.#####")
    mantStr = Replace(mantStr, " ", "")

    Dim tpSci As String
    tpSci = mantStr & "\cdot10^{" & CStr(exp) & "}"

    Dim containsTp As Boolean
    containsTp = (InStr(1, norm, tpPlain, vbTextCompare) > 0) Or _
                 (InStr(1, norm, tpSci, vbTextCompare) > 0)

    If hasLam Then
        AssertTrue containsTp, "SubstituteFailure must contain numeric tp when lambdas exist"
    Else
        AssertTrue Not containsTp, "SubstituteFailure must NOT contain numeric tp when no lambdas exist"
    End If

    Debug.Print "OK: SubstituteFailure tp presence matches lambda presence for f=" & fName
    Exit Sub

EH:
    Debug.Print "FAIL: Test_SubstituteFailure_Includes_tp_numeric_only_when_lambdas_exist: " & Err.Description
    MsgBox "FAIL: Test_SubstituteFailure_Includes_tp_numeric_only_when_lambdas_exist" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub Test_SingleExternStageQ_DoesNotPrintWi_AndDoesNotUseTp()
    On Error GoTo EH

    Dim fName As String, extName As String
    If Not FindSingleExternFunction(fName, extName) Then
        Debug.Print "SKIP: Test_SingleExternStageQ_DoesNotPrintWi_AndDoesNotUseTp (no suitable single-extern function found)"
        Exit Sub
    End If

    InitGlobals

    Dim extID As Long
    extID = GetID(extName)

    If Not IsExternID(extID) Then
        Debug.Print "SKIP: extern not loaded? (" & extName & ")"
        Exit Sub
    End If

    Dim qi As Object
    Set qi = m_ExternByID(extID)

    If Not CBool(qi("HasStages")) Then
        Debug.Print "SKIP: extern '" & extName & "' has no per-stage values"
        Exit Sub
    End If

    Dim sSym As String
    sSym = RewriteFailure(fName, 0)

    AssertTrue Not ContainsTpSymbolic(sSym), "Single stage-Q term must not contain tp"
    AssertTrue InStr(1, sSym, "W_{", vbTextCompare) = 0, "Single stage-Q term must not contain Wi"

    Dim sNum As String
    sNum = SubstituteFailure(fName, 0)
    AssertTrue Not ContainsTpSymbolic(sNum), "Single stage-Q numeric must not contain tp"

    Debug.Print "OK: Single per-stage extern Q skips Wi and tp for function '" & fName & "' (extern '" & extName & "')"
    Exit Sub

EH:
    Debug.Print "FAIL: Test_SingleExternStageQ_DoesNotPrintWi_AndDoesNotUseTp: " & Err.Description
    MsgBox "FAIL: Test_SingleExternStageQ_DoesNotPrintWi_AndDoesNotUseTp" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub RunAll_Core_Tests()
    Debug.Print "=== RUN CORE TESTS ==="
    Test_Tp_Affects_CalcFailure_By_LambdaCount
    
    Dim tpl As Object: Set tpl = LoadFormatTemplates()

Debug.Print "SYM_Q_TEMPLATE = "; GetTplWarn(tpl, "SYM_Q_TEMPLATE", "<missing>")
Debug.Print "TP_SYM_POW     = "; GetTplWarn(tpl, "TP_SYM_POW", "<missing>")
Debug.Print "SYM_MULT_TEMPLATE = "; GetTplWarn(tpl, "SYM_MULT_TEMPLATE", "<missing>")
Debug.Print "SYM_TERM_TEMPLATE = "; GetTplWarn(tpl, "SYM_TERM_TEMPLATE", "<missing>")
    
    Test_RewriteFailure_Includes_tp_only_when_lambdas_exist
    Test_SubstituteFailure_Includes_tp_numeric_only_when_lambdas_exist
    Test_SingleExternStageQ_DoesNotPrintWi_AndDoesNotUseTp
    Debug.Print "=== DONE CORE TESTS ==="
End Sub


Public Sub MinTest_ApplyTokens()
    Debug.Print "--- MinTest_ApplyTokens ---"
    Dim tpl As String
    tpl = "A=[[A]]; B=[[B]]; C={{latex}}"
    Dim out As String
    out = ApplyTokens(tpl, Array("A", "B"), Array("1", "2"))
    Debug.Print out
    If out <> "A=1; B=2; C={{latex}}" Then
        Err.Raise vbObjectError + 9101, "MinTest", "ApplyTokens работает не так, как ожидается: " & out
    End If
End Sub


