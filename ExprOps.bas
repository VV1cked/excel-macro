Attribute VB_Name = "ExprOps"
Option Explicit

' ================================
' Модуль ExprOps – операции над логическими выражениями
' ================================


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


' --- AND выражение: перемножаем термы ---
Public Function MultiplyExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
    Dim res As New CExpr, t1() As CTerm, t2() As CTerm
    Dim i As Long, j As Long, newTerm As CTerm
    
    t1 = e1.GetTerms(): t2 = e2.GetTerms()
    If (Not Not t1) = 0 Or (Not Not t2) = 0 Then
        Set MultiplyExpr = res
        Exit Function
    End If
    
    For i = LBound(t1) To UBound(t1)
        For j = LBound(t2) To UBound(t2)
            Set newTerm = FastMerge(t1(i), t2(j))
            If Not newTerm Is Nothing Then
                ' Ограничение по порядку – теперь по реальному rTerm, а не по количеству факторов
                Dim rTerm As Long
                rTerm = TermTotalOrderFromIDs(newTerm.FactorIDs)
                
                If rTerm <= R_MAX Then
                    res.AddTerm newTerm
                End If
            End If
        Next j
    Next i
    
    Set MultiplyExpr = res
End Function


' --- Быстрое объединение термов ---
Public Function FastMerge(ByVal t1 As CTerm, ByVal t2 As CTerm) As CTerm
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim id As Variant, resIDs() As Long, keys As Variant, i As Long
    
    For Each id In t1.FactorIDs: dict(id) = Empty: Next
    For Each id In t2.FactorIDs: dict(id) = Empty: Next
    
    If dict.Count = 0 Then
        Set FastMerge = Nothing
        Exit Function
    End If
    
    ReDim resIDs(0 To dict.Count - 1)
    keys = dict.keys
    For i = 0 To UBound(keys): resIDs(i) = CLng(keys(i)): Next
    
    SortIDs resIDs
    
    Dim sKeys() As String: ReDim sKeys(UBound(resIDs))
    For i = 0 To UBound(resIDs): sKeys(i) = CStr(resIDs(i)): Next
    
    Set FastMerge = New CTerm
    FastMerge.Init resIDs, t1.Multiplier * t2.Multiplier, Join(sKeys, "|")
End Function


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


' ====== CalcExpr ======
' Теперь единая точка расчёта: используем новый CalcExprFailure (с Q-подсистемами и корректным tp/Wi)
Public Function CalcExpr(ByVal e As CExpr, Optional ByVal stage As Variant = 0) As Double
    CalcExpr = CalcExprFailure(e, stage)
End Function


' ====== Helpers ======
' Полный порядок терма по списку ID факторов:
'   r = (#lambda) + sum(orderQ)
' где orderQ берём из m_ExternByID(id)("Order"), если фактор является внешней подсистемой.
Private Function TermTotalOrderFromIDs(ByRef ids() As Long) As Long
    Dim j As Long
    Dim nLambda As Long: nLambda = 0
    Dim sumRQ As Long: sumRQ = 0
    
    For j = LBound(ids) To UBound(ids)
        Dim id As Long: id = ids(j)
        If Not m_ExternByID Is Nothing Then
            If m_ExternByID.Exists(id) Then
                Dim qi As Object: Set qi = m_ExternByID(id)
                sumRQ = sumRQ + CLng(qi("Order"))
            Else
                nLambda = nLambda + 1
            End If
        Else
            nLambda = nLambda + 1
        End If
    Next j
    
    TermTotalOrderFromIDs = nLambda + sumRQ
End Function

