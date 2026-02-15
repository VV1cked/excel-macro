Attribute VB_Name = "ExprOps"
Option Explicit

' ================================
' ExprOps v2 - Операции над выражениями с компактными термами
' ================================

' --- OR выражение: объединяем термы ---
Public Function OrExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
    Dim res As New CExpr
    Dim tArr() As CTerm
    Dim i As Long
    
    ' Добавляем термы из e1
    If Not e1 Is Nothing Then
        tArr = e1.GetTerms()
        
        ' ===== ИСПРАВЛЕНИЕ: Безопасная проверка =====
        On Error Resume Next
        Dim ub1 As Long, lb1 As Long
        lb1 = LBound(tArr)
        ub1 = UBound(tArr)
        On Error GoTo 0
        
        If Err.Number = 0 And ub1 >= lb1 Then
            For i = lb1 To ub1
                On Error Resume Next
                Dim tempTerm As CTerm
                Set tempTerm = tArr(i)
                If Not tempTerm Is Nothing Then
                    res.AddTerm tempTerm
                End If
            Next i
        End If
    End If
    
    ' Добавляем термы из e2
    If Not e2 Is Nothing Then
        tArr = e2.GetTerms()
        
        On Error Resume Next
        Dim ub2 As Long, lb2 As Long
        lb2 = LBound(tArr)
        ub2 = UBound(tArr)
        On Error GoTo 0
        
        If Err.Number = 0 And ub2 >= lb2 Then
            For i = lb2 To ub2
                On Error Resume Next
                Dim tempTerm2 As CTerm
                Set tempTerm2 = tArr(i)
                If Not tempTerm2 Is Nothing Then
                    res.AddTerm tempTerm2
                End If
            Next i
        End If
    End If
    
    Set OrExpr = res
End Function

' --- AND выражение: перемножаем термы ---
Public Function MultiplyExpr(ByVal e1 As CExpr, ByVal e2 As CExpr) As CExpr
    Dim res As New CExpr
    Dim t1() As CTerm, t2() As CTerm
    Dim v1 As Variant, v2 As Variant
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long

    ' --- защита от Nothing ---
    If e1 Is Nothing Or e2 Is Nothing Then
        Set MultiplyExpr = res
        Exit Function
    End If

    t1 = e1.GetTerms()
    t2 = e2.GetTerms()

    v1 = t1
    v2 = t2

    If Not TryGetBounds(v1, lb1, ub1) Then
        Set MultiplyExpr = res
        Exit Function
    End If

    If Not TryGetBounds(v2, lb2, ub2) Then
        Set MultiplyExpr = res
        Exit Function
    End If

    ' --- проверка первых термов на Nothing ---
    If t1(lb1) Is Nothing Or t2(lb2) Is Nothing Then
        Set MultiplyExpr = res
        Exit Function
    End If

    ' ====== ПРОВЕРКА КОМПАКТНОСТИ ======
    Dim order1 As Long, order2 As Long
    Dim uniform1 As Boolean, uniform2 As Boolean
    Dim canCompact As Boolean

    uniform1 = IsUniformOrderSum(e1, order1)
    uniform2 = IsUniformOrderSum(e2, order2)
    canCompact = uniform1 And uniform2

    If canCompact Then
        Dim totalOrder As Long
        totalOrder = order1 + order2

        If totalOrder <= R_MAX Then
            Dim newTerm As CTerm
            Set newTerm = CreateCompactOrHybridTerm(e1, e2, totalOrder)
            If Not newTerm Is Nothing Then res.AddTerm newTerm
        End If

        Set MultiplyExpr = res
        Exit Function
    End If

    ' ====== ОБЫЧНОЕ РАСКРЫТИЕ ======
    Dim i As Long, j As Long
    Dim merged As CTerm
    For i = lb1 To ub1
        If Not t1(i) Is Nothing Then
            For j = lb2 To ub2
                If Not t2(j) Is Nothing Then
                    Set merged = FastMerge(t1(i), t2(j))
                    If Not merged Is Nothing Then
                        Dim rTerm As Long
                        rTerm = TermTotalOrderFromIDs(merged.FactorIDs)
                        If rTerm <= R_MAX Then res.AddTerm merged
                    End If
                End If
            Next j
        End If
    Next i

    Set MultiplyExpr = res
End Function


' ====== НОВОЕ: Проверка унифицированности Order ======
' Возвращает True, если все термы в expr имеют одинаковый Order
' outOrder - этот единый Order
Private Function IsUniformOrderSum(ByVal expr As CExpr, ByRef outOrder As Long) As Boolean
    On Error GoTo ErrHandler

    If expr Is Nothing Then
        IsUniformOrderSum = False
        Exit Function
    End If

    Dim terms() As CTerm
    terms = expr.GetTerms()

    Dim vTerms As Variant
    vTerms = terms

    Dim lb As Long, ub As Long
    If Not TryGetBounds(vTerms, lb, ub) Then
        IsUniformOrderSum = False
        Exit Function
    End If

    If terms(lb) Is Nothing Then
        IsUniformOrderSum = False
        Exit Function
    End If

    Dim firstOrder As Long
    Dim orderSet As Boolean
    orderSet = False

    Dim i As Long
    For i = lb To ub
        If terms(i) Is Nothing Then
            IsUniformOrderSum = False
            Exit Function
        End If

        Dim termOrder As Long

        Select Case terms(i).TermType
            Case 1 ' ttCompact (если у тебя константа другая — замени на ttCompact)
                termOrder = terms(i).Order

            Case ttCachedFunc
                Dim orderVec As Object
                Set orderVec = GetOrComputeOrderVector(terms(i).FuncName)

                If orderVec Is Nothing Or orderVec.Count <> 1 Then
                    IsUniformOrderSum = False
                    Exit Function
                End If

                Dim ks As Variant
                ks = orderVec.keys
                termOrder = CLng(ks(0))

            Case Else ' ttNormal
                Dim ids() As Long
                ids = terms(i).FactorIDs

                Dim vIDs As Variant
                vIDs = ids

                Dim idLb As Long, idUb As Long
                If Not TryGetBounds(vIDs, idLb, idUb) Then
                    IsUniformOrderSum = False
                    Exit Function
                End If

                ' 1 ID? возможно это функция через ID->Name
                If idUb = idLb Then
                    Dim id As Long: id = ids(idLb)
                    Dim nm As String: nm = vbNullString
                    If id > 0 And id <= UBound(m_IDToName) Then nm = m_IDToName(id)

                    If Len(nm) > 0 And Not m_NameKind Is Nothing And m_NameKind.Exists(nm) And m_NameKind(nm) = "FUNC" Then
                        Dim ov As Object
                        Set ov = GetOrComputeOrderVector(nm)

                        If ov Is Nothing Or ov.Count <> 1 Then
                            IsUniformOrderSum = False
                            Exit Function
                        End If

                        ks = ov.keys
                        termOrder = CLng(ks(0))
                    Else
                        termOrder = TermTotalOrderFromIDs(ids)
                    End If
                Else
                    termOrder = TermTotalOrderFromIDs(ids)
                End If
        End Select

        If Not orderSet Then
            firstOrder = termOrder
            orderSet = True
        Else
            If termOrder <> firstOrder Then
                IsUniformOrderSum = False
                Exit Function
            End If
        End If
    Next i

    outOrder = firstOrder
    IsUniformOrderSum = True
    Exit Function

ErrHandler:
    Debug.Print "IsUniformOrderSum Error #" & Err.Number & ": " & Err.Description
    IsUniformOrderSum = False
End Function

' ====== НОВОЕ: Создание компактного/гибридного терма со склейкой факторов ======
Private Function CreateCompactOrHybridTerm(ByVal e1 As CExpr, ByVal e2 As CExpr, ByVal totalOrder As Long) As CTerm
    Dim factors As New Collection
    
    ' Извлекаем факторы из e1 (с разворачиванием вложенных компактных термов)
    ExtractFactors e1, factors
    
    ' Извлекаем факторы из e2
    ExtractFactors e2, factors
    
    ' Создаём компактный терм
    Set CreateCompactOrHybridTerm = New CTerm
    CreateCompactOrHybridTerm.InitCompact factors, 1#
    CreateCompactOrHybridTerm.Order = totalOrder
End Function

' Извлечение факторов из выражения (со склейкой)
Private Sub ExtractFactors(ByVal expr As CExpr, ByRef factors As Collection)
    If expr Is Nothing Then
        ' Если внезапно Nothing — добавим как "пустой" фактор? Обычно лучше просто выйти.
        Exit Sub
    End If

    Dim terms() As CTerm
    terms = expr.GetTerms()

    Dim vTerms As Variant
    vTerms = terms

    Dim lb As Long, ub As Long
    If TryGetBounds(vTerms, lb, ub) Then
        If ub = lb Then
            If Not terms(lb) Is Nothing Then
                If terms(lb).IsCompact Then
                    Dim f As Variant
                    For Each f In terms(lb).CompactFactors
                        factors.Add f
                    Next f
                    Exit Sub
                End If
            End If
        End If
    End If

    ' Иначе добавляем всё выражение как один фактор
    factors.Add expr
End Sub

' ====== ИСПРАВЛЕННЫЙ FastMerge: сохранение кратностей ======
' Операции НЕ логические: a1*a1 = a1^2, а не a1
Public Function FastMerge(ByRef t1 As CTerm, ByRef t2 As CTerm) As CTerm
    Dim ids1() As Long
    Dim ids2() As Long

    ids1 = t1.FactorIDs
    ids2 = t2.FactorIDs

    Dim v1 As Variant, v2 As Variant
    v1 = ids1
    v2 = ids2

    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    If Not TryGetBounds(v1, lb1, ub1) Then Exit Function
    If Not TryGetBounds(v2, lb2, ub2) Then Exit Function

    Dim n1 As Long, n2 As Long
    n1 = ub1 - lb1 + 1
    n2 = ub2 - lb2 + 1

    Dim resIDs() As Long
    ReDim resIDs(0 To n1 + n2 - 1)

    Dim i As Long, pos As Long
    pos = 0

    For i = lb1 To ub1
        resIDs(pos) = ids1(i)
        pos = pos + 1
    Next i

    For i = lb2 To ub2
        resIDs(pos) = ids2(i)
        pos = pos + 1
    Next i

    SortIDs resIDs

    Dim sKeys() As String
    ReDim sKeys(0 To UBound(resIDs))
    For i = 0 To UBound(resIDs)
        sKeys(i) = CStr(resIDs(i))
    Next i

    Set FastMerge = New CTerm
    FastMerge.Init resIDs, t1.Multiplier * t2.Multiplier, Join(sKeys, "|")
End Function


' Сортировка массива ID (простая пузырьковая сортировка)
Private Sub SortIDs(ByRef arr() As Long)
    Dim i As Long, j As Long, temp As Long
    Dim n As Long
    n = UBound(arr)
    
    For i = 0 To n - 1
        For j = i + 1 To n
            If arr(j) < arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' ====== Helpers ======
Public Function TermTotalOrderFromIDs(ByRef ids() As Long) As Long
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

