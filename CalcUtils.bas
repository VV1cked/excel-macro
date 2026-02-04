Attribute VB_Name = "CalcUtils"
Option Explicit

'' ===== Вычисление выражения =====
'
'Public Function CalcExpr(ByVal e As CExpr, Optional ByVal Stage As Variant = 0) As Double
'    Dim t() As CTerm
'    t = e.GetTerms()
'    If (Not Not t) = 0 Then Exit Function
'
'    Dim i As Long, j As Long, total As Double, p As Double
'    Dim orderIdx As Long
'    Dim currentIDs() As Long
'
'
'    Dim stageIdx As Long
'    If IsNumeric(Stage) Then
'        stageIdx = CLng(Stage)
'    Else
'        stageIdx = -1 ' Если Stage="ALL", можно использовать stageIdx=0 или Wi=1
'    End If
'
'
'    ' Отладочный вывод
'    For i = LBound(t) To UBound(t)
'        currentIDs = t(i).FactorIDs
'
'        ' Проверяем пустой массив перед Join
'        Dim idsText As String
'        If (Not Not currentIDs) = 0 Or UBound(currentIDs) < LBound(currentIDs) Then
'            idsText = "(пусто)"
'        Else
'            Dim k As Long
'            idsText = ""
'            For k = LBound(currentIDs) To UBound(currentIDs)
'                idsText = idsText & currentIDs(k)
'                If k < UBound(currentIDs) Then idsText = idsText & ","
'            Next k
'        End If
'
'        Debug.Print "Term=" & t(i).Key & ", Order=" & t(i).Order & _
'                    ", FactorIDs=" & idsText & ", Multiplier=" & t(i).Multiplier
'    Next i
'
'    ' Основной расчёт
'    For i = LBound(t) To UBound(t)
'        p = 1#
'        orderIdx = t(i).Order
'        currentIDs = t(i).FactorIDs
'
'        ' Преумножаем ?
'        If (Not Not currentIDs) <> 0 And UBound(currentIDs) >= LBound(currentIDs) Then
'            For j = LBound(currentIDs) To UBound(currentIDs)
'                p = p * m_LambdaValues(currentIDs(j))
'            Next j
'        End If
'
'
'
'        ' Умножаем на Wi
'        If orderIdx <= R_MAX Then
'            Dim wi As Double
'            If stageIdx = -1 Then
'                ' Stage="ALL" > используем Wi=1
'                wi = 1#
'            Else
'                wi = m_WiValues(orderIdx, stageIdx)
'            End If
'
'            total = total + (CDbl(t(i).Multiplier) * wi * p)
'        End If
'    Next i
'
'    CalcExpr = total
'End Function





'Public Function CalcExpr(ByVal e As CExpr, Optional ByVal Stage As Long = 0) As Double
'    Dim t() As CTerm: t = e.GetTerms()
'    If (Not Not t) = 0 Then
'        Debug.Print "CalcExpr: Нет термов"
'        Exit Function
'    End If
'
'    Dim i As Long, j As Long, total As Double, p As Double
'    Dim orderIdx As Long
'    Dim currentIDs() As Long
'
'    For i = LBound(t) To UBound(t)
'        ' Проверка, что t(i) не Nothing
'        If t(i) Is Nothing Then
'            Debug.Print "CalcExpr: t(" & i & ") = Nothing"
'            GoTo NextTerm
'        End If
'
'        currentIDs = t(i).FactorIDs
'        ' Проверка, что массив не пуст
'        If (Not Not currentIDs) = 0 Then
'            Debug.Print "CalcExpr: t(" & i & ").FactorIDs пустой"
'            GoTo NextTerm
'        End If
'
'        Debug.Print "Term=" & t(i).Key & ", Order=" & t(i).Order & _
'                    ", FactorIDs=" & Join(currentIDs, ",")
'NextTerm:
'    Next i
'
'    ' Основной расчёт
'    For i = LBound(t) To UBound(t)
'        If t(i) Is Nothing Then GoTo NextCalc
'        currentIDs = t(i).FactorIDs
'        If (Not Not currentIDs) = 0 Then GoTo NextCalc
'
'        p = 1#
'        orderIdx = t(i).Order
'        For j = LBound(currentIDs) To UBound(currentIDs)
'            If currentIDs(j) > UBound(m_LambdaValues) Then
'                Debug.Print "OutOfRange: LambdaValues(" & currentIDs(j) & ")"
'                p = 0
'            Else
'                p = p * m_LambdaValues(currentIDs(j))
'            End If
'        Next j
'
'        If orderIdx <= R_MAX Then
'            total = total + (CDbl(t(i).Multiplier) * m_WiValues(orderIdx) * p)
'        End If
'NextCalc:
'    Next i
'
'    CalcExpr = total
'End Function


' ====== Преобразование строки в число с поддержкой "e" и локализации ======
Public Function ParseDouble(ByVal v As Variant, ByVal context As String) As Double
    Dim s As String
    s = Trim(CStr(v))
    
    ' Убираем пробелы
    s = Replace(s, " ", "")
    ' Заменяем запятую на точку для десятичных
    s = Replace(s, ",", ".")
    
    ' Попытка через Val (работает с экспонентой)
    If s <> "" Then
        ParseDouble = Val(s)
    Else
        Err.Raise 996, , "Невозможно преобразовать в число: " & context & " = '" & v & "'"
    End If
End Function

