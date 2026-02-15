Attribute VB_Name = "CalcFailureMain"
Option Explicit

' ====== Global caches / mappings ======
Public m_IDToName() As String
Public m_LambdaValues() As Double
Public m_NameToID As Object

Public m_FuncExprCache As Object
Public m_FuncDNFCache As Object
Public m_CallStack As Object

Public m_WiValues() As Double   ' (r, stage)
Public m_Tp As Double
Public m_NameKind As Object  ' name -> "ELEM" | "FUNC" | "Q"
Public m_FuncOrderVectorCache As Object

' ====== External (precalculated) subsystem Q data ======
' Keyed by subsystem ID (Long) -> Dictionary with fields:
'   "Name" (String)
'   "Order" (Long)                     ' default 1 if blank
'   "HasStages" (Boolean)              ' True if user provided 13 values
'   "QAll" (Double)                    ' Q for all time (tp) (for stages: Sum(QStage))
'   "QStage" (Variant array)           ' 0..12 when HasStages=True
Public m_ExternByID As Object

' Sheet name for external systems
Private Const SHEET_EXTERN As String = "ExternSystems"

'=========================================================
' Public API
'=========================================================

Public Function CalcFailure(ByVal FuncName As String, Optional ByVal stage As Variant = 0) As Double
    On Error GoTo ErrHandler

    InitGlobals
    m_CallStack.RemoveAll

    Dim e As CExpr
    Set e = EvalFunction(Trim$(FuncName))

    If e Is Nothing Then
        CalcFailure = 0#
        Exit Function
    End If

    CalcFailure = CalcExprFailure(e, stage)
    Exit Function

ErrHandler:
    MsgBox "Ошибка расчёта функции '" & FuncName & "': " & Err.Description, vbCritical
    CalcFailure = 0#
End Function

'=========================================================
' Initialization
'=========================================================

Public Sub InitGlobals()
    Set m_NameToID = CreateObject("Scripting.Dictionary")
    Set m_FuncExprCache = CreateObject("Scripting.Dictionary")
    Set m_FuncDNFCache = CreateObject("Scripting.Dictionary")
    Set m_CallStack = CreateObject("Scripting.Dictionary")
    Set m_ExternByID = CreateObject("Scripting.Dictionary")
    Set m_NameKind = CreateObject("Scripting.Dictionary")
    Set m_FuncOrderVectorCache = CreateObject("Scripting.Dictionary")

    ReDim m_IDToName(0)
    ReDim m_LambdaValues(0)

    LoadLambdas
    LoadFunctions
    LoadWi
    LoadTp
    LoadExternSystems
End Sub

Private Sub RegisterName(ByVal nm As String, ByVal kind As String, ByVal where As String)
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Sub

    If m_NameKind.Exists(nm) Then
        Dim prev As String: prev = CStr(m_NameKind(nm))
        Err.Raise vbObjectError + 3201, "InitGlobals", _
            "Конфликт имён: '" & nm & "' уже занято (" & prev & "), нельзя зарегистрировать как " & kind & _
            ". Источник: " & where
    End If

    m_NameKind.Add nm, kind
End Sub

'=========================================================
' Load tp from Elements!C (first positive numeric)
'=========================================================

Public Sub LoadTp()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_ELEMENTS)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim v As Variant
        v = ws.Cells(r, 3).value

        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                m_Tp = CDbl(v)
                Exit Sub
            End If
        End If
    Next r

    Err.Raise 996, , "Не найден tp на листе " & SHEET_ELEMENTS & " (колонка C)"
End Sub

'=========================================================
' Load lambdas from Elements sheet
'=========================================================

Public Sub LoadLambdas()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_ELEMENTS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_ELEMENTS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).value

    Dim i As Long, id As Long, sName As String
    For i = 1 To UBound(data, 1)
        sName = Trim$(CStr(data(i, RANGE_ELEMENTS_COL_NAME)))
        If sName <> "" Then
            ' 1) запрет дублей элементов
            If m_NameKind.Exists(sName) Then
                Err.Raise vbObjectError + 3202, "LoadLambdas", "Дублируется элемент '" & sName & "' на листе " & SHEET_ELEMENTS
            End If

            ' 2) регистрируем как элемент
            Call RegisterName(sName, "ELEM", SHEET_ELEMENTS & "!" & "A" & (i + 1))

            ' 3) теперь можно создавать ID
            id = GetID(sName)
            If id > UBound(m_LambdaValues) Then ReDim Preserve m_LambdaValues(0 To id + 50)
            m_LambdaValues(id) = ParseDouble(CStr(data(i, RANGE_ELEMENTS_COL_LAMBDA)), sName)
        End If
Next i
End Sub

'=========================================================
' Load Functions cache
'=========================================================

Public Sub LoadFunctions()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_FUNCTIONS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_FUNCTIONS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).value

    Dim i As Long, fName As String
    For i = 1 To UBound(data, 1)
        fName = Trim$(CStr(data(i, RANGE_FUNCTIONS_COL_NAME)))
        If fName <> "" Then
            ' нельзя совпадать с элементом
            If m_NameKind.Exists(fName) Then
                Err.Raise vbObjectError + 3203, "LoadFunctions", _
                "Имя функции '" & fName & "' конфликтует с ранее загруженным именем (" & m_NameKind(fName) & "). Лист: " & SHEET_FUNCTIONS
            End If

            Call RegisterName(fName, "FUNC", SHEET_FUNCTIONS & "!" & "A" & (i + 1))
            m_FuncExprCache(fName) = Trim$(CStr(data(i, RANGE_FUNCTIONS_COL_EXPR)))
        End If
    Next i
End Sub

'=========================================================
' Load Wi table (dynamic max r, fixed stages 0..12)
'=========================================================

Public Sub LoadWi()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_WI)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, WI_COL_R).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range(WI_COL_R & "2:" & WI_COL_MAX & lastRow).value

    Dim i As Long, rIdx As Long, maxR As Long
    maxR = R_MAX
    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) Then
            rIdx = CLng(data(i, 1))
            If rIdx > maxR Then maxR = rIdx
        End If
    Next i

    ReDim m_WiValues(0 To maxR, 0 To 12)

    Dim stage As Long
    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) Then
            rIdx = CLng(data(i, 1))
            If rIdx >= 0 And rIdx <= maxR Then
                For stage = 0 To 12
                    m_WiValues(rIdx, stage) = ParseDouble(data(i, stage + 2), "Wi r=" & rIdx & " stage=" & stage)
                Next stage
            End If
        End If
    Next i
End Sub

'=========================================================
' Load external subsystem Q from sheet ExternSystems
'=========================================================

Private Sub LoadExternSystems()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_EXTERN)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    For r = 2 To lastRow
        Dim nm As String
        nm = Trim$(CStr(ws.Cells(r, 1).value))
        If Len(nm) = 0 Then GoTo NextRow

        Dim qCell As Variant
        qCell = ws.Cells(r, 2).value

        Dim ordCell As Variant
        ordCell = ws.Cells(r, 3).value

        Dim ord As Long
        If IsNumeric(ordCell) Then
            ord = CLng(ordCell)
            If ord <= 0 Then ord = 1
        Else
            ord = 1
        End If

        Dim qInfo As Object
        Set qInfo = ParseExternQCell(qCell, nm)

        qInfo("Name") = nm
        qInfo("Order") = ord

        If m_NameKind.Exists(nm) Then
            Err.Raise vbObjectError + 3204, "LoadExternSystems", _
            "Имя внешней Q '" & nm & "' конфликтует с ранее загруженным именем (" & m_NameKind(nm) & "). Лист: " & SHEET_EXTERN & " строка " & r
        End If

        Call RegisterName(nm, "Q", SHEET_EXTERN & "!" & "A" & r)

        Dim id As Long
        id = GetID(nm)

        If m_ExternByID.Exists(id) Then
            Set m_ExternByID(id) = qInfo
        Else
            m_ExternByID.Add id, qInfo   ' Add accepts object as Variant
        End If

NextRow:
    Next r
End Sub

Private Function ParseExternQCell(ByVal v As Variant, ByVal contextName As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim s As String
    s = Trim$(CStr(v))

    If Len(s) = 0 Then
        Err.Raise vbObjectError + 750, , "ExternSystems: пустое поле вероятности для '" & contextName & "'"
    End If

    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ";", " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Dim parts() As String
    parts = Split(s, " ")

    Dim nums() As Double
    Dim n As Long, i As Long
    n = 0
    ReDim nums(0 To 0)

    For i = LBound(parts) To UBound(parts)
        Dim tok As String
        tok = Trim$(parts(i))
        If Len(tok) > 0 Then
            If n = 0 Then
                nums(0) = ParseDouble(tok, "ExternSystems Q '" & contextName & "'")
            Else
                ReDim Preserve nums(0 To n)
                nums(n) = ParseDouble(tok, "ExternSystems Q '" & contextName & "'")
            End If
            n = n + 1
        End If
    Next i

    If n = 1 Then
        d("HasStages") = False
        d("QAll") = nums(0)

    ElseIf n = 13 Then
        d("HasStages") = True

        Dim qStage As Variant
        ReDim qStage(0 To 12)   ' IMPORTANT: Variant array for Dictionary storage

        Dim sumAll As Double
        sumAll = 0#

        For i = 0 To 12
            qStage(i) = CDbl(nums(i))
            sumAll = sumAll + CDbl(nums(i))
        Next i

        d("QStage") = qStage
        d("QAll") = sumAll

    Else
        Err.Raise vbObjectError + 751, , "ExternSystems: для '" & contextName & "' должно быть 1 или 13 чисел, найдено: " & CStr(n)
    End If

    Set ParseExternQCell = d
End Function

'=========================================================
' ID mapping
'=========================================================


Public Function GetID(ByVal sName As String) As Long
    Dim newID As Long
    sName = Trim$(sName)

    If Not m_NameToID.Exists(sName) Then
        newID = m_NameToID.Count + 1
        m_NameToID(sName) = newID

        If newID > UBound(m_IDToName) Then ReDim Preserve m_IDToName(0 To newID + 50)
        m_IDToName(newID) = sName

        GetID = newID
    Else
        GetID = m_NameToID(sName)
    End If
End Function


Public Function GetIDStrict(ByVal sName As String, Optional ByVal ctx As String = "") As Long
    sName = Trim$(sName)
    If Len(sName) = 0 Then
        Err.Raise vbObjectError + 1001, "Parser", "Пустое имя атома. " & ctx
    End If

    If m_NameKind Is Nothing Then
        Err.Raise vbObjectError + 3999, "Resolver", "m_NameKind не инициализирован. " & ctx
    End If

    If Not m_NameKind.Exists(sName) Then
        Err.Raise vbObjectError + 3002, "Resolver", "Неизвестное имя в формуле: '" & sName & "'. " & ctx
    End If

    ' Теперь безопасно: имя гарантированно "из таблиц", GetID не создаст мусор (оно уже будет в m_NameToID)
    GetIDStrict = GetID(sName)
End Function


' ================================
' Расширения CalcFailureMain v2
' ================================
' Добавить в начало CalcFailureMain.bas после объявления глобальных переменных:
'
' Public m_FuncOrderVectorCache As Object  ' FuncName ? Dictionary(Order ? Value)
'
' Добавить в InitGlobals после Set m_NameKind:
'
' Set m_FuncOrderVectorCache = CreateObject("Scripting.Dictionary")
'

' ====== НОВОЕ: Вычисление и кеширование OrderVector функции ======
Public Function GetOrComputeOrderVector(ByVal fName As String) As Object
    On Error GoTo ErrHandler

    fName = Trim$(fName)

    ' Проверяем кеш
    If Not m_FuncOrderVectorCache Is Nothing Then
        If m_FuncOrderVectorCache.Exists(fName) Then
            Set GetOrComputeOrderVector = m_FuncOrderVectorCache(fName)
            Exit Function
        End If
    End If

    ' Если кеш ещё не инициализирован (на всякий случай)
    If m_FuncOrderVectorCache Is Nothing Then
        Set m_FuncOrderVectorCache = CreateObject("Scripting.Dictionary")
    End If

    ' Вычисляем функцию
    Dim expr As CExpr
    Set expr = EvalFunction(fName)

    Dim orderVec As Object
    Set orderVec = CreateObject("Scripting.Dictionary")

    If expr Is Nothing Then
        Set m_FuncOrderVectorCache(fName) = orderVec
        Set GetOrComputeOrderVector = orderVec
        Exit Function
    End If

    ' Получаем термы
    Dim terms() As CTerm
    terms = expr.GetTerms()

    ' --- Безопасно проверяем, что массив terms валиден и не пуст ---
    Dim vTerms As Variant
    vTerms = terms

    Dim lbT As Long, ubT As Long
    If Not TryGetBounds(vTerms, lbT, ubT) Then
        Set m_FuncOrderVectorCache(fName) = orderVec
        Set GetOrComputeOrderVector = orderVec
        Exit Function
    End If

    ' Группируем термы по Order
    Dim i As Long
    For i = lbT To ubT
        If Not terms(i) Is Nothing Then
            Dim r As Long
            Dim termValue As Double

            Select Case terms(i).TermType
                Case ttCompact
                    r = terms(i).Order
                    termValue = CalcCompactTerm(terms(i), 0, True)

                Case ttCachedFunc
                    ' Рекурсивно для вложенных функций
                    Dim nestedVec As Object
                    Set nestedVec = GetOrComputeOrderVector(terms(i).FuncName)

                    Dim k As Variant
                    For Each k In nestedVec.keys
                        If Not orderVec.Exists(k) Then orderVec(k) = 0#
                        orderVec(k) = CDbl(orderVec(k)) + CDbl(nestedVec(k)) * terms(i).Multiplier
                    Next k

                    GoTo NextTerm

                Case Else  ' ttNormal
                    r = TermTotalOrderFromIDs(terms(i).FactorIDs)
                    termValue = CalcSingleTerm(terms(i), 0, True)
            End Select

            If Not orderVec.Exists(r) Then orderVec(r) = 0#
            orderVec(r) = CDbl(orderVec(r)) + termValue
        End If

NextTerm:
    Next i

    ' Кешируем
    Set m_FuncOrderVectorCache(fName) = orderVec
    Set GetOrComputeOrderVector = orderVec
    Exit Function

ErrHandler:
    Debug.Print "GetOrComputeOrderVector error: " & Err.Number & " - " & Err.Description & " (Func=" & fName & ")"
    ' В случае ошибки возвращаем пустой вектор, но не падаем
    Dim emptyVec As Object
    Set emptyVec = CreateObject("Scripting.Dictionary")
    If Not m_FuncOrderVectorCache Is Nothing Then
        On Error Resume Next
        Set m_FuncOrderVectorCache(fName) = emptyVec
        On Error GoTo 0
    End If
    Set GetOrComputeOrderVector = emptyVec
End Function

' ====== НОВОЕ: Вычисление компактного терма ======
Public Function CalcCompactTerm(ByRef t As CTerm, ByVal st As Long, ByVal isAll As Boolean) As Double
    On Error GoTo ErrHandler
    Dim r As Long
    r = t.Order
    
    Dim wiValue As Double
    If Not isAll Then
        wiValue = GetWiSafe(r, st)
    Else
        wiValue = 1#
    End If
    
    Dim tpPow As Double
    tpPow = m_Tp ^ r
    
    ' Вычисляем произведение факторов
    Dim factorProduct As Double
    factorProduct = 1#
    
    Dim f As Variant
    For Each f In t.CompactFactors
        Dim factorValue As Double
        
        Select Case TypeName(f)
            Case "CTerm"
                ' Кешированная функция
                Dim ft As CTerm
                Set ft = f
                
                If ft.TermType = ttCachedFunc Then
                    factorValue = ft.GetValueForOrder(r)
                Else
                    ' Ошибка архитектуры
                    factorValue = 0#
                End If
                
            Case "CExpr"
                ' Обычный CExpr фактор (сумма элементов)
                Dim fCExpr As CExpr
                Set fCExpr = f
                factorValue = CalcFactorSum(fCExpr, r, st, isAll)
                
            Case Else
                factorValue = 0#
        End Select
        
        factorProduct = factorProduct * factorValue
    Next f
    
    CalcCompactTerm = t.Multiplier * wiValue * tpPow * factorProduct
    Exit Function
ErrHandler:
    Debug.Print "CalcCompact " & Err.Description & " Order= " & r
    CalcCompactTerm = 0#
End Function

' Вычисление суммы факторов для заданного Order
Private Function CalcFactorSum(ByRef factorExpr As CExpr, ByVal expectedOrder As Long, ByVal st As Long, ByVal isAll As Boolean) As Double
    On Error GoTo ErrHandler

    Dim terms() As CTerm
    terms = factorExpr.GetTerms()

    ' --- безопасная проверка массива terms ---
    Dim vTerms As Variant
    vTerms = terms

    Dim lbT As Long, ubT As Long
    If Not TryGetBounds(vTerms, lbT, ubT) Then
        CalcFactorSum = 0#
        Exit Function
    End If

    Dim sumValue As Double
    sumValue = 0#

    Dim i As Long
    For i = lbT To ubT
        If Not terms(i) Is Nothing Then

            ' В factorExpr ожидаем "атомарные" термы: один ID (элемент или функция)
            Dim ids() As Long
            ids = terms(i).FactorIDs

            ' --- безопасная проверка массива ids ---
            Dim vIDs As Variant
            vIDs = ids

            Dim lbI As Long, ubI As Long
            If Not TryGetBounds(vIDs, lbI, ubI) Then
                GoTo NextTerm
            End If

            ' Только один фактор (atom)
            If ubI <> lbI Then GoTo NextTerm

            Dim id As Long
            id = ids(lbI)

            ' ID -> Name
            Dim nm As String
            nm = vbNullString
            If id > 0 Then
                If id <= UBound(m_IDToName) Then nm = m_IDToName(id)
            End If
            If Len(nm) = 0 Then GoTo NextTerm

            ' Функция или элемент?
            If Not m_NameKind Is Nothing Then
                If m_NameKind.Exists(nm) Then
                    If CStr(m_NameKind(nm)) = "FUNC" Then
                        ' --- ФУНКЦИЯ ---
                        Dim funcExpr As CExpr
                        Set funcExpr = EvalFunction(nm)

                        Dim funcVal As Double
                        If isAll Then
                            ' Если у тебя реально есть режим "ALL", оставляем как было.
                            ' Иначе можно заменить на 0/стадию по умолчанию.
                            funcVal = CalcExprFailure(funcExpr, "ALL")
                        Else
                            funcVal = CalcExprFailure(funcExpr, st)
                        End If

                        sumValue = sumValue + funcVal * terms(i).Multiplier
                        GoTo NextTerm
                    End If
                End If
            End If

            ' --- ЭЛЕМЕНТ ---
            If id > 0 Then
                If id <= UBound(m_LambdaValues) Then
                    sumValue = sumValue + m_LambdaValues(id) * terms(i).Multiplier
                End If
            End If

        End If

NextTerm:
    Next i

    CalcFactorSum = sumValue
    Exit Function

ErrHandler:
    Debug.Print "CalcFactorSum error: " & Err.Number & " - " & Err.Description
    CalcFactorSum = 0#
End Function

' ====== МОДИФИЦИРОВАННОЕ: CalcExprFailure с поддержкой компактных термов ======
' Заменить существующую функцию CalcExprFailure в CalcFailureMain.bas
' ====== CalcExprFailure с поддержкой ttCompact и ttCachedFunc ======
Public Function CalcExprFailure(ByVal e As CExpr, Optional ByVal stage As Variant = 0) As Double
    On Error GoTo ErrHandler

    If e Is Nothing Then
        CalcExprFailure = 0#
        Exit Function
    End If

    Dim st As Long
    If IsNumeric(stage) Then
        st = CLng(stage)
    Else
        ' если передали "ALL" или что-то нечисловое — считаем как ALL
        st = 0
    End If

    Dim isAll As Boolean
    isAll = (Not IsNumeric(stage)) ' "ALL" / Variant(String) и т.п.

    Dim terms() As CTerm
    terms = e.GetTerms()

    ' --- безопасная проверка массива terms ---
    Dim vTerms As Variant
    vTerms = terms

    Dim lbT As Long, ubT As Long
    If Not TryGetBounds(vTerms, lbT, ubT) Then
        CalcExprFailure = 0#
        Exit Function
    End If

    Dim total As Double
    total = 0#

    Dim i As Long
    For i = lbT To ubT
        If Not terms(i) Is Nothing Then
            Select Case terms(i).TermType

                Case ttCompact
                    ' Компактный терм: в режиме ALL Wi=1
                    total = total + CalcCompactTerm(terms(i), st, isAll)

                Case ttCachedFunc
                    ' Кешированная функция: берём её OrderVector и подставляем Wi,tp
                    Dim orderVec As Object
                    Set orderVec = GetOrComputeOrderVector(terms(i).FuncName)

                    If Not orderVec Is Nothing Then
                        Dim k As Variant
                        For Each k In orderVec.keys
                            Dim r As Long
                            r = CLng(k)

                            Dim wiValue As Double
                            If isAll Then
                                wiValue = 1#
                            Else
                                wiValue = GetWiSafe(r, st)
                            End If

                            Dim tpPow As Double
                            tpPow = m_Tp ^ r

                            total = total + terms(i).Multiplier * wiValue * tpPow * CDbl(orderVec(k))
                        Next k
                    End If

                Case Else ' ttNormal
                    ' Обычный терм: в режиме ALL Wi=1
                    total = total + CalcSingleTerm(terms(i), st, isAll)

            End Select
        End If
    Next i

    CalcExprFailure = total
    Exit Function

ErrHandler:
    Debug.Print "CalcExprFailure error: " & Err.Number & " - " & Err.Description
    CalcExprFailure = 0#
End Function


' ====== ВСПОМОГАТЕЛЬНОЕ: Вычисление одного обычного терма ======
Private Function CalcSingleTerm(ByRef t As CTerm, ByVal st As Long, ByVal isAll As Boolean) As Double
    Dim r As Long
    r = t.Order
    
    Dim wiValue As Double
    If Not isAll Then
        wiValue = GetWiSafe(r, st)
    Else
        wiValue = 1#
    End If
    
    Dim tpPow As Double
    tpPow = m_Tp ^ r
    
    Dim product As Double
    product = 1#
    
    Dim ids() As Long
    ids = t.FactorIDs
    
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        Dim id As Long
        id = ids(i)
        
        If id <= UBound(m_LambdaValues) Then
            product = product * m_LambdaValues(id)
        End If
    Next i
    
    CalcSingleTerm = t.Multiplier * wiValue * tpPow * product
End Function

' ====== ВСПОМОГАТЕЛЬНОЕ: Безопасное получение Wi ======
Private Function GetWiSafe(ByVal r As Long, ByVal st As Long) As Double
    On Error Resume Next
    GetWiSafe = m_WiValues(r, st)
    If Err.Number <> 0 Then GetWiSafe = 1#
    On Error GoTo 0
End Function


