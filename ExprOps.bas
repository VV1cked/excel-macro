Attribute VB_Name = "ExprOps"
Option Explicit

' ================================
' Модуль ExprOps – операции над логическими выражениями
' ================================



' ====== Объединение выражений (логическое OR) ======

' --- OR выражение: объединяем термы ---

Public Function OrExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
    Dim res As New CExpr
    Dim tArr() As CTerm
    Dim i As Long
    
    ' добавляем термы из e1
    If Not e1 Is Nothing Then
        tArr = e1.GetTerms()
        If (Not Not tArr) <> 0 Then
            For i = LBound(tArr) To UBound(tArr)
                res.AddTerm tArr(i)
            Next i
        End If
    End If
    
    ' добавляем термы из e2
    If Not e2 Is Nothing Then
        tArr = e2.GetTerms()
        If (Not Not tArr) <> 0 Then
            For i = LBound(tArr) To UBound(tArr)
                res.AddTerm tArr(i)
            Next i
        End If
    End If
    
    Set OrExpr = res
End Function

'Public Function OrExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
'    Dim res As New CExpr
'    Dim arr() As CTerm
'    Dim i As Long
'
'    ' Добавляем термы из e1
'    If Not e1 Is Nothing Then
'        arr = e1.GetTerms()
'        If (Not Not arr) <> 0 Then
'            For i = LBound(arr) To UBound(arr)
'                res.AddTerm arr(i)
'            Next i
'        End If
'    End If
'
'    ' Добавляем термы из e2
'    If Not e2 Is Nothing Then
'        arr = e2.GetTerms()
'        If (Not Not arr) <> 0 Then
'            For i = LBound(arr) To UBound(arr)
'                res.AddTerm arr(i)
'            Next i
'        End If
'    End If
'
'    Set OrExpr = res
'End Function


' ====== Умножение выражений (логическое AND) ======

' --- AND выражение: перемножаем термы ---
Public Function MultiplyExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
    Dim res As New CExpr, t1() As CTerm, t2() As CTerm
    Dim i As Long, j As Long, newTerm As CTerm
    t1 = e1.GetTerms(): t2 = e2.GetTerms()
    If (Not Not t1) = 0 Or (Not Not t2) = 0 Then Set MultiplyExpr = res: Exit Function
    
    For i = LBound(t1) To UBound(t1)
        For j = LBound(t2) To UBound(t2)
            If t1(i).Order + t2(j).Order <= R_MAX Then
                Set newTerm = FastMerge(t1(i), t2(j))
                If Not newTerm Is Nothing Then res.AddTerm newTerm
            End If
        Next j
    Next i
    Set MultiplyExpr = res
End Function

'
'Public Function AndExpr(ByVal e1 As CExpr, ByVal e2 As CExpr, Optional ByVal R_MAX As Long = 4) As CExpr
'    Dim res As New CExpr
'    Dim t1() As CTerm, t2() As CTerm
'    Dim i As Long, j As Long, newTerm As CTerm
'
'    t1 = e1.GetTerms(): t2 = e2.GetTerms()
'
'    If (Not Not t1) = 0 Or (Not Not t2) = 0 Then
'        Set AndExpr = res
'        Exit Function
'    End If
'
'    ' Перемножаем все термы
'    For i = LBound(t1) To UBound(t1)
'        For j = LBound(t2) To UBound(t2)
'            If t1(i).Order + t2(j).Order <= R_MAX Then
'                Set newTerm = FastMerge(t1(i), t2(j))
'                If Not newTerm Is Nothing Then res.AddTerm newTerm
'            End If
'        Next j
'    Next i
'
'    Set AndExpr = res
'End Function

' ====== Быстрое объединение двух термов без дублирования ======
' --- Быстрое объединение термов ---
Public Function FastMerge(ByVal t1 As CTerm, ByVal t2 As CTerm) As CTerm
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim id As Variant, resIDs() As Long, keys As Variant, i As Long
    
    For Each id In t1.FactorIDs: dict(id) = Empty: Next
    For Each id In t2.FactorIDs: dict(id) = Empty: Next
    
    ReDim resIDs(0 To dict.Count - 1)
    keys = dict.keys
    For i = 0 To UBound(keys): resIDs(i) = CLng(keys(i)): Next
    
    SortIDs resIDs
    
    Dim sKeys() As String: ReDim sKeys(UBound(resIDs))
    For i = 0 To UBound(resIDs): sKeys(i) = CStr(resIDs(i)): Next
    
    Set FastMerge = New CTerm
    FastMerge.Init resIDs, t1.Multiplier * t2.Multiplier, Join(sKeys, "|")
End Function

'Public Function FastMerge(ByVal t1 As CTerm, ByVal t2 As CTerm) As CTerm
'    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
'    Dim id As Variant
'
'    ' Собираем уникальные FactorIDs
'    For Each id In t1.FactorIDs: dict(id) = Empty: Next
'    For Each id In t2.FactorIDs: dict(id) = Empty: Next
'
'    ' Копируем в массив
'    Dim resIDs() As Long
'    ReDim resIDs(0 To dict.Count - 1)
'    Dim keys As Variant: keys = dict.keys
'    Dim i As Long
'    For i = 0 To UBound(keys)
'        resIDs(i) = CLng(keys(i))
'    Next
'
'    ' Сортировка ID для унификации ключа
'    Call SortIDs(resIDs)
'
'    ' Формируем строковый ключ
'    Dim sKeys() As String
'    ReDim sKeys(0 To UBound(resIDs))
'    For i = 0 To UBound(resIDs): sKeys(i) = CStr(resIDs(i)): Next
'
'    ' Создаем новый терм
'    Set FastMerge = New CTerm
'    FastMerge.Init resIDs, t1.Multiplier * t2.Multiplier, Join(sKeys, "|")
'End Function

' ====== Простая сортировка массива чисел ======
Public Sub SortIDs(ByRef arr() As Long)
    Dim i As Long, j As Long, temp As Long
    For i = LBound(arr) + 1 To UBound(arr)
        temp = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If arr(j) > temp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = temp
    Next i
End Sub

' ====== Объединение термов из source в target ======
Public Sub MergeExpr(ByRef target As CExpr, ByVal source As CExpr)
    Dim t() As CTerm
    t = source.GetTerms()
    
    If (Not Not t) = 0 Then Exit Sub
    
    Dim i As Long
    For i = LBound(t) To UBound(t)
        target.AddTerm t(i)
    Next i
End Sub

Public Function CalcExpr(ByVal e As CExpr, Optional ByVal stage As Variant = 0) As Double
    ' e         - CExpr для расчета
    ' Stage     - номер этапа: 0..12, или "ALL"

    Dim t() As CTerm
    t = e.GetTerms()
    If (Not Not t) = 0 Then Exit Function

    Dim i As Long, j As Long
    Dim total As Double, p As Double
    Dim orderIdx As Long
    Dim currentIDs() As Long
    Dim wiValue As Double
    Dim tpPow As Double

    For i = LBound(t) To UBound(t)
        p = 1#
        orderIdx = t(i).Order
        currentIDs = t(i).FactorIDs

        ' Перемножаем ? для всех элементов терма
        For j = LBound(currentIDs) To UBound(currentIDs)
            p = p * m_LambdaValues(currentIDs(j))
        Next j

        ' Учитываем время: домножаем терм на tp^order
        If orderIdx > 0 Then
            tpPow = m_Tp ^ orderIdx
        Else
            tpPow = 1#
        End If

        ' Выбираем Wi в зависимости от Stage
        If stage = "ALL" Then
            wiValue = 1#
        Else
            If orderIdx <= R_MAX Then
                wiValue = m_WiValues(orderIdx, CLng(stage))
            Else
                wiValue = 0#
            End If
        End If

        ' Вклад терма
        total = total + (CDbl(t(i).Multiplier) * wiValue * p * tpPow)
    Next i

    CalcExpr = total
End Function


'Public Function CalcExpr(ByVal e As CExpr, Optional ByVal Stage As Variant = 0) As Double
'    Dim t() As CTerm
'    t = e.GetTerms()
'    If (Not Not t) = 0 Then Exit Function
'
'    Dim i As Long, j As Long, total As Double, p As Double
'    Dim orderIdx As Long
'    Dim currentIDs() As Long
'    Dim stageIdx As Long
'
'    ' Stage=ALL > wi=1
'    If VarType(Stage) = vbString Then
'        If UCase(Stage) = "ALL" Then stageIdx = -1 Else stageIdx = CLng(Stage)
'    Else
'        stageIdx = CLng(Stage)
'    End If
'
'    For i = LBound(t) To UBound(t)
'        currentIDs = t(i).FactorIDs
'
'        ' Рассчитываем произведение ?
'        p = 1#
'        If (Not Not currentIDs) <> 0 And UBound(currentIDs) >= LBound(currentIDs) Then
'            For j = LBound(currentIDs) To UBound(currentIDs)
'                p = p * m_LambdaValues(currentIDs(j))
'            Next j
'        End If
'
'        orderIdx = t(i).Order
'        Dim wi As Double
'        If stageIdx = -1 Then
'            wi = 1#
'        Else
'            wi = m_WiValues(orderIdx, stageIdx)
'        End If
'
'        total = total + t(i).Multiplier * wi * p
'    Next i
'
'    CalcExpr = total
'End Function

