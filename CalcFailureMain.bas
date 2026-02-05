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
        v = ws.Cells(r, 3).Value

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

Private Sub LoadLambdas()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_ELEMENTS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_ELEMENTS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value

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

Private Sub LoadFunctions()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_FUNCTIONS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_FUNCTIONS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value

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

Private Sub LoadWi()
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_WI)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, WI_COL_R).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range(WI_COL_R & "2:" & WI_COL_MAX & lastRow).Value

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
        nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nm) = 0 Then GoTo NextRow

        Dim qCell As Variant
        qCell = ws.Cells(r, 2).Value

        Dim ordCell As Variant
        ordCell = ws.Cells(r, 3).Value

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



'=========================================================
' Failure calculator aware of external Q subsystems
'=========================================================

Public Function CalcExprFailure(ByVal e As CExpr, Optional ByVal stage As Variant = 0) As Double
    Dim t() As CTerm
    t = e.GetTerms()
    If (Not Not t) = 0 Then Exit Function

    Dim total As Double: total = 0#

    Dim isAll As Boolean
    isAll = (VarType(stage) = vbString And UCase$(CStr(stage)) = "ALL")

    Dim st As Long
    If Not isAll Then
        st = CLng(stage)
        If st < 0 Or st > 12 Then Err.Raise vbObjectError + 760, , "Stage вне диапазона 0..12: " & CStr(stage)
    Else
        st = 0
    End If

    Dim i As Long
    For i = LBound(t) To UBound(t)
        If t(i) Is Nothing Then GoTo NextTerm
        If Abs(t(i).Multiplier) < 0.0000000001 Then GoTo NextTerm

        ' IMPORTANT: FactorIDs -> Variant (avoid error 450)
        Dim idsV As Variant
        idsV = t(i).FactorIDs
        If IsEmpty(idsV) Then GoTo NextTerm

        Dim factorCount As Long
        factorCount = UBound(idsV) - LBound(idsV) + 1
        If factorCount <= 0 Then GoTo NextTerm

        Dim nLambda As Long: nLambda = 0
        Dim sumRQ As Long: sumRQ = 0
        Dim lambdaProd As Double: lambdaProd = 1#
        Dim hasQ As Boolean: hasQ = False

        Dim j As Long
        For j = LBound(idsV) To UBound(idsV)
            Dim id As Long: id = CLng(idsV(j))

            If Not m_ExternByID Is Nothing And m_ExternByID.Exists(id) Then
                hasQ = True
                Dim qi As Object: Set qi = m_ExternByID(id)
                sumRQ = sumRQ + CLng(qi("Order"))
            Else
                nLambda = nLambda + 1
                lambdaProd = lambdaProd * m_LambdaValues(id)
            End If
        Next j

        Dim rTerm As Long
        rTerm = nLambda + sumRQ

        Dim onlyQ As Boolean
        onlyQ = (hasQ And nLambda = 0 And factorCount = 1)

        Dim skipWi As Boolean: skipWi = False
        Dim qPart As Double: qPart = 1#

        If onlyQ Then
            Dim singleQId As Long
            singleQId = CLng(idsV(LBound(idsV)))
            Dim singleQInfo As Object: Set singleQInfo = m_ExternByID(singleQId)

            If CBool(singleQInfo("HasStages")) Then
                skipWi = True
                If isAll Then
                    qPart = CDbl(singleQInfo("QAll"))
                Else
                    Dim arrV As Variant
                    arrV = singleQInfo("QStage")
                    qPart = CDbl(arrV(st))
                End If
            Else
                qPart = CDbl(singleQInfo("QAll"))
            End If

        Else
            If hasQ Then
                Dim qProd As Double: qProd = 1#
                For j = LBound(idsV) To UBound(idsV)
                    Dim id2 As Long: id2 = CLng(idsV(j))
                    If m_ExternByID.Exists(id2) Then
                        Dim qi2 As Object: Set qi2 = m_ExternByID(id2)
                        qProd = qProd * CDbl(qi2("QAll"))
                    End If
                Next j
                qPart = qProd
            Else
                qPart = 1#
            End If
        End If

        Dim tpPow As Double
        If nLambda > 0 Then tpPow = m_Tp ^ nLambda Else tpPow = 1#

        Dim wiValue As Double: wiValue = 1#
        If (Not isAll) And (Not skipWi) Then
            wiValue = GetWiSafe(rTerm, st)
        End If

        total = total + (CDbl(t(i).Multiplier) * wiValue * lambdaProd * tpPow * qPart)

NextTerm:
    Next i

    CalcExprFailure = total
End Function

Private Function GetWiSafe(ByVal r As Long, ByVal st As Long) As Double
    Dim maxR As Long
    maxR = UBound(m_WiValues, 1)

    If r < 0 Or r > maxR Then
        Err.Raise vbObjectError + 761, , "Wi: порядок r=" & CStr(r) & " вне диапазона 0.." & CStr(maxR) & ". Проверь таблицу Wi."
    End If

    GetWiSafe = m_WiValues(r, st)
End Function

