Attribute VB_Name = "Tests"
Option Explicit

Sub Test_FailureCalc()
    Dim result As Double
    Dim testFuncname As String
        
    testFuncname = "SYS6"
    result = CalcFailure(testFuncname, "ALL")
    
End Sub

Sub Test_FullCheck()
    On Error Resume Next
    Set m_NameToID = Nothing
    Set m_FuncExprCache = Nothing
    Set m_FuncDNFCache = Nothing
    
    Dim ws As Worksheet
    Set ws = Sheets("Functions")
    If ws Is Nothing Then
        MsgBox "Ошибка листа функц"
    End If
    
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Debug.Print "Строка " & i & ": [" & ws.Cells(i, 1).Value & "]"
    Next i
    Dim testName As String
    testName = Trim(ws.Cells(2, 1).Value)
    If testName = "" Then
        Exit Sub
    End If
    
    Debug.Print "Запуск CalcFailure для: '" & testName & "'"
    
    Debug.Print "Результат: " & CalcFailure(testName)
End Sub

Sub TestCalcFailure_SYS6()
    Dim result As Double
    Dim FuncName As String
    Dim stage As String
    
    FuncName = "SYS6"   ' имя новой функции на листе Functions
    stage = "0"         ' можно пробовать "0", "1", "ALL"
    
    On Error GoTo ErrHandler
    
    Debug.Print "=== Начало теста CalcFailure для " & FuncName & ", Stage=" & stage & " ==="
    
    ' Вызов функции
    result = CalcFailure(FuncName, stage)
    
    Debug.Print "Результат CalcFailure(" & FuncName & ", Stage=" & stage & ") = " & result
    Debug.Print "=== Конец теста ==="
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Ошибка: " & Err.Number & " - " & Err.Description
    MsgBox "Ошибка в тесте CalcFailure: " & Err.Description, vbCritical
End Sub

Sub TestSYS6_Terms()
    ' Инициализация глобальных кешей
    InitGlobals
    
    Dim e As CExpr
    ' Теперь EvalFunction безопасно вызываем
    Set e = EvalFunction("SYS6")  ' Stage = 0
    
    If e Is Nothing Then
        Debug.Print "EvalFunction вернула Nothing"
        Exit Sub
    End If
    
    Dim t() As CTerm
    t = e.GetTerms()
    
    If (Not Not t) = 0 Then
        Debug.Print "EvalFunction вернула CExpr без термов!"
        Exit Sub
    End If
    
    Dim i As Long, ids() As Long
    For i = LBound(t) To UBound(t)
        ids = t(i).FactorIDs
        Dim idsText As String
        
        ' Проверяем, что массив инициализирован и содержит хотя бы один элемент
        If (Not Not ids) = 0 Or UBound(ids) < LBound(ids) Then
            idsText = "(пусто)"
        Else
            ' Конвертируем элементы массива в строку
            Dim j As Long
            idsText = ""
            For j = LBound(ids) To UBound(ids)
                idsText = idsText & ids(j)
                If j < UBound(ids) Then idsText = idsText & ","
            Next j
        End If
        
        Debug.Print "Term=" & t(i).key & ", Order=" & t(i).Order & _
                    ", FactorIDs=" & idsText & ", Multiplier=" & t(i).Multiplier
    Next
        
    ' Попробуем сразу расчёт
    Dim result As Double
    result = CalcExpr(e, 0)
    Debug.Print "CalcExpr(SYS6, Stage=0) = " & result
End Sub

Public Sub TestCalcFailureStepByStep()
    Dim fSYS4 As Double, fSYS5 As Double, fSYS6 As Double
    Dim e As CExpr
    Dim t() As CTerm
    Dim i As Long, j As Long, ids() As Long
    Dim idsText As String
    
    Debug.Print "=== Начало теста CalcFailure ==="
    
    ' 0. Инициализация кешей и данных
    InitGlobals
    
    ' 1. Рассчитываем SYS4
    fSYS4 = CalcFailure("SYS4", 3)
    Debug.Print "SYS4 = " & fSYS4
    
    ' 2. Рассчитываем SYS5
    fSYS5 = CalcFailure("SYS5", 3)
    Debug.Print "SYS5 = " & fSYS5
    
    ' 3. Рассчитываем SYS6 (SYS5*SYS4)
    fSYS6 = CalcFailure("SYS6", 3)
    Debug.Print "SYS6 = " & fSYS6
    
    ' 4. Разбираем термы SYS6
    Set e = EvalFunction("SYS6")
    t = e.GetTerms()
    
    Debug.Print "=== Термы SYS6 ==="
    For i = LBound(t) To UBound(t)
        ids = t(i).FactorIDs
        If (Not Not ids) = 0 Then
            idsText = "(пусто)"
        Else
            ' Join безопасно с Variant
            Dim vIDs() As Variant
            ReDim vIDs(LBound(ids) To UBound(ids))
            For j = LBound(ids) To UBound(ids): vIDs(j) = ids(j): Next j
            idsText = Join(vIDs, ",")
        End If
        
        Debug.Print "Term " & i & ": Key=" & t(i).key & _
                    ", Order=" & t(i).Order & _
                    ", Multiplier=" & t(i).Multiplier & _
                    ", FactorIDs=" & idsText
    Next i
    
    ' 5. Проверка произведения вручную
    Debug.Print "Проверка: SYS5 * SYS4 = " & fSYS5 * fSYS4
    
    Debug.Print "=== Конец теста CalcFailure ==="
End Sub


Public Sub TestCalcFailureStepByStep_MultiStage()
    Dim stages As Variant
    stages = Array(0, 3, 12, "ALL")
    
    Dim s As Variant
    Dim fSYS4 As Double, fSYS5 As Double, fSYS6 As Double
    Dim eSYS6 As CExpr
    Dim t() As CTerm
    Dim i As Long, j As Long, ids() As Long
    Dim idsText As String
    
    Debug.Print "=== Начало теста CalcFailure с несколькими этапами ==="
    
    ' Инициализация кешей и данных
    InitGlobals
    
    For Each s In stages
        Debug.Print ">>> Этап: " & s
        
        ' Вычисляем функции через CalcFailure
        fSYS4 = CalcFailure("SYS4", s)
        fSYS5 = CalcFailure("SYS5", s)
        fSYS6 = CalcFailure("SYS6", s)
        
        Debug.Print "SYS4 = " & fSYS4
        Debug.Print "SYS5 = " & fSYS5
        Debug.Print "SYS6 = " & fSYS6
        
        ' Разбираем термы SYS6
        Set eSYS6 = EvalFunction("SYS6")
        t = eSYS6.GetTerms()
        
        Debug.Print "--- Термы SYS6 ---"
        For i = LBound(t) To UBound(t)
            ids = t(i).FactorIDs
            If (Not Not ids) = 0 Then
                idsText = "(пусто)"
            Else
                ' Join безопасно с Variant
                Dim vIDs() As Variant
                ReDim vIDs(LBound(ids) To UBound(ids))
                For j = LBound(ids) To UBound(ids): vIDs(j) = ids(j): Next j
                idsText = Join(vIDs, ",")
            End If
            
            Debug.Print "Term " & i & ": Key=" & t(i).key & _
                        ", Order=" & t(i).Order & _
                        ", Multiplier=" & t(i).Multiplier & _
                        ", FactorIDs=" & idsText
        Next i
        
        ' Проверка произведения SYS5*SYS4
        Debug.Print "Проверка: SYS5 * SYS4 = " & fSYS5 * fSYS4
        Debug.Print "---------------------------------------"
    Next s
    
    Debug.Print "=== Конец теста CalcFailure ==="
End Sub


'===========================
' Helpers for tests
'===========================

Private Function FirstFunctionName() As String
    ' Берём первую функцию из листа Functions (A2)
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
    ' Находит первую числовую ячейку в Elements!C начиная со 2-й строки
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

    Err.Raise vbObjectError + 901, , "Не удалось найти числовую ячейку tp на листе Elements в колонке C"
End Function


Private Sub SetTpAndReinit(ByVal tp As Double)
    Dim c As Range
    Set c = FirstTpCell()
    c.Value = tp

    ' Перезагружаем кэши и tp
    InitGlobals
    m_CallStack.RemoveAll
End Sub


Private Function ContainsLatexToken(ByVal s As String, ByVal token As String) As Boolean
    ContainsLatexToken = (InStr(1, s, token, vbTextCompare) > 0)
End Function


Private Sub AssertTrue(ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then Err.Raise vbObjectError + 910, , "ASSERT TRUE FAILED: " & msg
End Sub

Private Sub AssertNear(ByVal a As Double, ByVal b As Double, ByVal relTol As Double, ByVal msg As String)
    ' относительная погрешность
    If a = 0 And b = 0 Then Exit Sub
    Dim denom As Double
    denom = IIf(Abs(b) > Abs(a), Abs(b), Abs(a))
    If denom = 0 Then Err.Raise vbObjectError + 911, , "ASSERT NEAR FAILED (both near 0?): " & msg

    If Abs(a - b) / denom > relTol Then
        Err.Raise vbObjectError + 912, , msg & " (a=" & a & ", b=" & b & ", rel=" & Abs(a - b) / denom & ")"
    End If
End Sub


'===========================
' Tests
'===========================

Public Sub Test_Tp_Affects_CalcFailure_By_Order()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    ' Берём выражение и определяем порядок первого терма (r)
    InitGlobals
    Dim e As CExpr
    Set e = EvalFunction(fName)

    Dim terms() As CTerm
    terms = e.GetTerms()
    If (Not Not terms) = 0 Then Err.Raise vbObjectError + 920, , "Функция '" & fName & "' дала пустой набор термов"

    Dim r As Long
    r = terms(LBound(terms)).Order
    AssertTrue r > 0, "Order первого терма должен быть > 0"

    ' Выбираем стадию 0, чтобы было проще/стабильнее
    Dim stage As Variant
    stage = 0

    ' Считаем при двух tp
    Dim tp1 As Double, tp2 As Double
    tp1 = 0.5
    tp2 = 1#

    SetTpAndReinit tp1
    Dim v1 As Double
    v1 = CalcFailure(fName, stage)

    SetTpAndReinit tp2
    Dim v2 As Double
    v2 = CalcFailure(fName, stage)

    ' При неизменных ? и Wi отношение должно быть примерно (tp1/tp2)^r
    Dim expected As Double
    expected = (tp1 / tp2) ^ r

    ' v1 ~= v2 * expected
    AssertNear v1, v2 * expected, 0.000001, "CalcFailure не масштабируется как tp^r (f=" & fName & ", r=" & r & ")"

    Debug.Print "OK: CalcFailure scales by tp^r for f=" & fName & ", r=" & r
    Exit Sub

EH:
    Debug.Print "FAIL: Test_Tp_Affects_CalcFailure_By_Order: " & Err.Description
    MsgBox "FAIL: Test_Tp_Affects_CalcFailure_By_Order" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub Test_RewriteFailure_Includes_tp_power()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    SetTpAndReinit 0.5

    Dim s As String
    s = RewriteFailure(fName, 0)

    ' Строка должна быть валидным LaTeX и начинаться с Q_{name}
    AssertTrue Left$(s, 3) = "Q_{", "RewriteFailure должен начинаться с 'Q_{'"

    ' Должно быть упоминание t_p (либо t_p, либо t_p^{...})
    AssertTrue ContainsLatexToken(s, "t_p"), "В RewriteFailure должен присутствовать t_p"

    Debug.Print "OK: RewriteFailure contains t_p for f=" & fName
    Exit Sub

EH:
    Debug.Print "FAIL: Test_RewriteFailure_Includes_tp_power: " & Err.Description
    MsgBox "FAIL: Test_RewriteFailure_Includes_tp_power" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub Test_SubstituteFailure_Includes_tp_numeric_power()
    On Error GoTo EH

    Dim fName As String
    fName = FirstFunctionName()

    Dim tp As Double
    tp = 0.5

    SetTpAndReinit tp

    Dim s As String
    s = SubstituteFailure(fName, 0)

    Debug.Print "SubstituteFailure output:"
    Debug.Print s

    AssertTrue Left$(s, 3) = "Q_{", "SubstituteFailure должен начинаться с 'Q_{'"

    ' Нормализуем пробелы (шаблоны могут добавлять пробелы внутри степеней)
    Dim norm As String
    norm = Replace(s, " ", "")

    ' Вариант 1: plain (0,5)
    Dim tpPlain As String
    tpPlain = Format$(tp, "0.############") ' в RU-локали будет "0,5"
    tpPlain = Replace(tpPlain, " ", "")

    ' Вариант 2: scientific (5\cdot 10^{-1})
    Dim av As Double, exp As Long, mant As Double
    av = Abs(tp)
    exp = Fix(Log(av) / Log(10#))
    mant = tp / (10# ^ exp)

    Dim mantStr As String
    mantStr = Format$(mant, "0.#####") ' для 0.5 будет "5"
    mantStr = Replace(mantStr, " ", "")

    Dim tpSci As String
    tpSci = mantStr & "\cdot10^{" & CStr(exp) & "}" ' без пробелов
    ' В формуле может быть "\cdot 10^{...}" — убираем пробелы в norm, поэтому так ок

    Dim ok As Boolean
    ok = (InStr(1, norm, tpPlain, vbTextCompare) > 0) Or _
         (InStr(1, norm, tpSci, vbTextCompare) > 0)

    AssertTrue ok, "В SubstituteFailure должно присутствовать tp либо в plain (" & tpPlain & "), либо в scientific (" & tpSci & ")"

    Debug.Print "OK: SubstituteFailure contains numeric tp for f=" & fName
    Exit Sub

EH:
    Debug.Print "FAIL: Test_SubstituteFailure_Includes_tp_numeric_power: " & Err.Description
    MsgBox "FAIL: Test_SubstituteFailure_Includes_tp_numeric_power" & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub RunAll_Tp_Tests()
    Debug.Print "=== RUN TP TESTS ==="
    Test_Tp_Affects_CalcFailure_By_Order
    Test_RewriteFailure_Includes_tp_power
    Test_SubstituteFailure_Includes_tp_numeric_power
    Debug.Print "=== DONE TP TESTS ==="
End Sub

