Attribute VB_Name = "CalcFailureMain"
Option Explicit
' ====== Глобальные кэши ======
'Public m_Codes() As Integer
Public m_IDToName() As String
Public m_LambdaValues() As Double
Public m_NameToID As Object
Public m_FuncExprCache As Object
Public m_FuncDNFCache As Object
Public m_CallStack As Object
Public m_WiValues() As Double
Public m_Tp As Double

' ====== Инициализация и точка входа ======

Public Function CalcFailure(ByVal FuncName As String, Optional ByVal stage As Variant = 0) As Double
    On Error GoTo ErrHandler
    
    InitGlobals
    m_CallStack.RemoveAll
    
    Dim e As CExpr
    Set e = EvalFunction(Trim(FuncName))
    If e Is Nothing Then
        CalcFailure = 0#
        Exit Function
    End If
    
    ' Рассчитываем выражение с учетом этапа
    CalcFailure = CalcExpr(e, stage)
    Exit Function
    
ErrHandler:
    MsgBox "Ошибка расчета функции '" & FuncName & "': " & Err.Description, vbCritical
    CalcFailure = 0#
End Function


'Public Function CalcFailure(ByVal FuncName As String, Optional ByVal Stage As String = "ALL") As Double
'    On Error GoTo ErrHandler
'    InitGlobals
'    m_CallStack.RemoveAll
'
'    Dim e As CExpr
'    Set e = EvalFunction(Trim(FuncName))
'    If e Is Nothing Then
'        CalcFailure = 0#: Exit Function
'    End If
'
'    Dim useAllWi As Boolean
'    Dim stageNum As Long
'
'    Stage = UCase(Trim(Stage))
'    If Stage = "ALL" Then
'        useAllWi = True
'    Else
'        If IsNumeric(Stage) Then
'            stageNum = CLng(Stage)
'            If stageNum < 0 Or stageNum > 12 Then
'                Err.Raise 998, , "Недопустимый Stage: " & Stage
'            End If
'        Else
'            Err.Raise 998, , "Недопустимый Stage: " & Stage
'        End If
'    End If
'
'    CalcFailure = CalcExprWithStage(e, useAllWi, stageNum)
'    Exit Function
'ErrHandler:
'    MsgBox "Ошибка расчета '" & FuncName & "' : " & Err.Description, vbCritical
'    CalcFailure = 0#
'End Function

Private Sub EnsureGlobals()
    If m_NameToID Is Nothing Then Set m_NameToID = CreateObject("Scripting.Dictionary")
    If m_FuncExprCache Is Nothing Then Set m_FuncExprCache = CreateObject("Scripting.Dictionary")
    If m_FuncDNFCache Is Nothing Then Set m_FuncDNFCache = CreateObject("Scripting.Dictionary")
    If m_CallStack Is Nothing Then Set m_CallStack = CreateObject("Scripting.Dictionary")
End Sub


Public Function CalcExprWithStage(ByVal e As CExpr, ByVal useAllWi As Boolean, ByVal stageNum As Long) As Double
    Dim t() As CTerm, i As Long, j As Long, total As Double, p As Double
    Dim orderIdx As Long, currentIDs() As Long
    
    t = e.GetTerms()
    If (Not Not t) = 0 Then Exit Function
    
    total = 0#
    
    For i = LBound(t) To UBound(t)
        p = 1#
        orderIdx = t(i).Order
        currentIDs = t(i).FactorIDs
        
        ' Множитель ?
        For j = LBound(currentIDs) To UBound(currentIDs)
            p = p * m_LambdaValues(currentIDs(j))
        Next j
        
        ' Множитель Wi
        If useAllWi Then
            p = p * 1#
        Else
            If orderIdx <= R_MAX Then
                p = p * m_WiValues(orderIdx, stageNum)
            End If
        End If
        
        ' Умножаем на Multiplier терма
        total = total + p * t(i).Multiplier
    Next i
    
    CalcExprWithStage = total
End Function



'Public Function CalcFailure(ByVal FuncName As String, Optional ByVal Stage As Long = 0) As Double
'    On Error GoTo ErrHandler
'
'    InitGlobals
'    m_CallStack.RemoveAll
'
'    Dim e As CExpr
'    Set e = EvalFunction(Trim(FuncName))
'    If e Is Nothing Then
'        CalcFailure = 0#: Exit Function
'    End If
'
'    CalcFailure = CalcExpr(e, Stage)
'    Exit Function

'ErrHandler:
'    MsgBox "Ошибка расчета '" & FuncName & "' : " & Err.Description, vbCritical
'    CalcFailure = 0#
'End Function

' ====== Инициализация глобальных кэшей ======
Public Sub InitGlobals()
    Set m_NameToID = CreateObject("Scripting.Dictionary")
    Set m_FuncExprCache = CreateObject("Scripting.Dictionary")
    Set m_FuncDNFCache = CreateObject("Scripting.Dictionary")
    Set m_CallStack = CreateObject("Scripting.Dictionary")
    
    
    ReDim m_IDToName(0)
    ReDim m_LambdaValues(0)
    ReDim m_WiValues(0 To R_MAX)
    
    
    LoadLambdas
    LoadFunctions
    LoadWi
    LoadTp
End Sub

Public Sub LoadTp()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Elements")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim v As Variant
        v = ws.Cells(r, 3).Value ' column C

        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                m_Tp = CDbl(v)
                Exit Sub
            End If
        End If
    Next r

    Err.Raise 996, , "Не найдено tp на листе Elements (колонка C)"
End Sub

' ====== Загрузка значений элементов ======
Private Sub LoadLambdas()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_ELEMENTS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_ELEMENTS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value
    
    Dim i As Long, id As Long, sName As String
    For i = 1 To UBound(data, 1)
        sName = Trim(CStr(data(i, RANGE_ELEMENTS_COL_NAME)))
        If sName <> "" Then
            id = GetID(sName)
            If id > UBound(m_LambdaValues) Then ReDim Preserve m_LambdaValues(0 To id + 50)
            m_LambdaValues(id) = ParseDouble(CStr(data(i, RANGE_ELEMENTS_COL_LAMBDA)), sName)
        End If
    Next i
End Sub

' ====== Загрузка функций ======
Private Sub LoadFunctions()
    Dim ws As Worksheet: Set ws = Sheets(SHEET_FUNCTIONS)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_FUNCTIONS_COL_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value
    
    Dim i As Long, fName As String
    For i = 1 To UBound(data, 1)
        fName = Trim(CStr(data(i, RANGE_FUNCTIONS_COL_NAME)))
        If fName <> "" Then m_FuncExprCache(fName) = Trim(CStr(data(i, RANGE_FUNCTIONS_COL_EXPR)))
    Next i
End Sub

'====== Загрузка Wi ======
Private Sub LoadWi()
    Dim data As Variant, i As Long, rIdx As Long, lastRow As Long
    Dim stage As Long, colOffset As Long
    
    lastRow = Sheets(SHEET_WI).Cells(Rows.Count, WI_COL_R).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    data = Sheets(SHEET_WI).Range(WI_COL_R & "2:" & WI_COL_MAX & lastRow).Value
    
    ' Обнуляем массив Wi (0..R_MAX, 0..12)
    ReDim m_WiValues(0 To R_MAX, 0 To 12)
    
    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) Then
            rIdx = CLng(data(i, 1))
            If rIdx >= 0 And rIdx <= R_MAX Then
                For stage = 0 To 12
                    ' Столбцы Stage0..Stage12 идут с 2 по 14
                    m_WiValues(rIdx, stage) = ParseDouble(data(i, stage + 2), "Wi r=" & rIdx & " stage=" & stage)
                Next stage
            End If
        End If
    Next i
End Sub





' ====== Загрузка Wi ======
'Private Sub LoadWi()
'    Dim ws As Worksheet: Set ws = Sheets(SHEET_WI)
'    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, RANGE_WI_COL_ORDER).End(xlUp).Row
'    If lastRow < 2 Then Exit Sub
'
'    Dim data As Variant
'    data = ws.Range("A2:B" & lastRow).Value
'
'    Dim i As Long, r As Long
'    For i = 1 To UBound(data, 1)
'        If IsNumeric(data(i, RANGE_WI_COL_ORDER)) Then
'            r = CLng(data(i, RANGE_WI_COL_ORDER))
'            If r >= 0 And r <= R_MAX Then
'                m_WiValues(r) = ParseDouble(CStr(data(i, RANGE_WI_COL_VALUE)), "Wi r=" & r)
'            End If
'        End If
'    Next i
'End Sub

' ====== Получение уникального ID для элемента/имени ======
Public Function GetID(ByVal sName As String) As Long
    Dim newID As Long
    sName = Trim(sName)
    
    If Not m_NameToID.Exists(sName) Then
        newID = m_NameToID.Count + 1
        m_NameToID(sName) = newID
        
        ' Расширяем массив ID > имя
        If newID > UBound(m_IDToName) Then ReDim Preserve m_IDToName(0 To newID + 50)
        m_IDToName(newID) = sName
        
        GetID = newID
    Else
        GetID = m_NameToID(sName)
    End If
End Function
