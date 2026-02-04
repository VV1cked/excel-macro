Attribute VB_Name = "Parser"
Option Explicit

'==============================
' Модуль: Parser + EvalFunction
'==============================

' Символы
Private Const CH_LPAREN As String = "("
Private Const CH_RPAREN As String = ")"
Private Const CH_PLUS As String = "+"
Private Const CH_MULT As String = "*"

'=========================================================
' Вычисление функции с учётом кэша и рекурсивных вызовов
'=========================================================

' --- Конвертирует строковое выражение функции в CExpr ---
Public Function EvalFunction(ByVal fName As String) As CExpr
    If Not m_FuncExprCache.Exists(fName) Then Err.Raise 998, , "Не найдена функция: " & fName
    If m_FuncDNFCache.Exists(fName) Then Set EvalFunction = m_FuncDNFCache(fName)
    If m_CallStack.Exists(fName) Then Err.Raise 997, , "Цикл: " & fName
    
    m_CallStack.Add fName, True
    Dim sExpr As String: sExpr = Replace(m_FuncExprCache(fName), " ", "")
    
    Dim res As CExpr
    Set res = ParseOr(sExpr)
    
    Set m_FuncDNFCache(fName) = res ' кэшируем только структуру
    m_CallStack.Remove fName
    Set EvalFunction = res
End Function

'Public Function EvalFunction(ByVal fName As String) As CExpr
'    ' Проверка существования функции
'    If Not m_FuncExprCache.Exists(fName) Then
'        Debug.Print "Функция не найдена: " & fName
'        Set EvalFunction = New CExpr
'        Exit Function
'    End If
'
'    ' Если уже вычислена и закэширована DNF
'    If m_FuncDNFCache.Exists(fName) Then
'        Set EvalFunction = m_FuncDNFCache(fName)
'        Exit Function
'    End If
'
'    ' Проверка на рекурсивный цикл
'    If m_CallStack.Exists(fName) Then
'        Debug.Print "Цикл функции: " & fName
'        Set EvalFunction = New CExpr
'        Exit Function
'    End If
'
'    ' Добавляем в стек вызова
'    m_CallStack.Add fName, True
'
'    ' Создаём локальный массив символов
'    Dim sExpr As String: sExpr = Replace(m_FuncExprCache(fName), " ", "")
'    Dim codes() As Long, i As Long
'    ReDim codes(1 To Len(sExpr))
'    For i = 1 To Len(sExpr)
'        codes(i) = AscW(Mid$(sExpr, i, 1))
'    Next i
'
'    ' Разбор выражения
'    Dim res As CExpr
'    Set res = ParseOr(codes, 1, UBound(codes))
'
'    ' Кэшируем результат
'    Set m_FuncDNFCache(fName) = res
'
'    ' Убираем из стека
'    m_CallStack.Remove fName
'
'    Set EvalFunction = res
'End Function

'==============================
' Парсер OR выражения
'==============================


' --- Парсинг выражения с OR (+) ---
Public Function ParseOr(ByVal s As String) As CExpr
    Dim res As New CExpr, parts As Collection, p As Variant
    Set parts = SplitTop(s, "+")
    For Each p In parts
        Set res = OrExpr(res, ParseAnd(p))
    Next p
    Set ParseOr = res
End Function



'Private Function ParseOr(ByRef codes() As Long, ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    Dim res As New CExpr
'    Dim lvl As Long: lvl = 0
'    Dim LastP As Long: LastP = sIdx
'    Dim i As Long
'
'    For i = sIdx To eIdx
'        Select Case codes(i)
'            Case AscW(CH_LPAREN): lvl = lvl + 1
'            Case AscW(CH_RPAREN): lvl = lvl - 1
'            Case AscW(CH_PLUS)
'                If lvl = 0 Then
'                    MergeExpr res, ParseAnd(codes, LastP, i - 1)
'                    LastP = i + 1
'                End If
'        End Select
'    Next i
'
'    MergeExpr res, ParseAnd(codes, LastP, eIdx)
'    Set ParseOr = res
'End Function

'==============================
' Парсер AND выражения
'==============================

' --- Парсинг AND (*) ---
Public Function ParseAnd(ByVal s As String) As CExpr
    Dim res As New CExpr, parts As Collection, p As Variant
    Set parts = SplitTop(s, "*")
    Dim first As Boolean: first = True
    For Each p In parts
        If first Then
            Set res = ParseFactor(p)
            first = False
        Else
            Set res = MultiplyExpr(res, ParseFactor(p))
        End If
    Next p
    Set ParseAnd = res
End Function


'Private Function ParseAnd(ByRef codes() As Long, ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    Dim res As CExpr, i As Long, lvl As Long, LastP As Long
'    Dim parts As New Collection
'
'    LastP = sIdx
'    lvl = 0
'    For i = sIdx To eIdx
'        Select Case codes(i)
'            Case AscW(CH_LPAREN): lvl = lvl + 1
'            Case AscW(CH_RPAREN): lvl = lvl - 1
'            Case AscW(CH_MULT)
'                If lvl = 0 Then
'                    parts.Add Array(LastP, i - 1)
'                    LastP = i + 1
'                End If
'        End Select
'    Next
'    parts.Add Array(LastP, eIdx)
'
'    ' Сначала парсим первый фактор
'    Set res = ParseFactor(codes, CLng(parts(1)(0)), CLng(parts(1)(1)))
'
'    ' Умножаем последующие факторы
'    Dim p As Long
'    For p = 2 To parts.Count
'        Set res = MultiplyExpr(res, ParseFactor(codes, CLng(parts(p)(0)), CLng(parts(p)(1))))
'    Next p
'
'    Set ParseAnd = res
'End Function

'==============================
' Парсер фактора (атом или скобки)
'==============================

Public Function ParseFactor(ByVal s As String) As CExpr
    If Left(s, 1) = "(" And Right(s, 1) = ")" Then
        Set ParseFactor = ParseOr(Mid(s, 2, Len(s) - 2))
    ElseIf m_FuncExprCache.Exists(s) Then
        Set ParseFactor = EvalFunction(s)
    Else
        Set ParseFactor = CreateAtom(s)
    End If
End Function


'Private Function ParseFactor(ByRef codes() As Long, ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    ' Скобки
'    If codes(sIdx) = AscW(CH_LPAREN) And codes(eIdx) = AscW(CH_RPAREN) Then
'        Set ParseFactor = ParseOr(codes, sIdx + 1, eIdx - 1)
'    Else
'        ' Собираем имя атома
'        Dim name As String: name = ""
'        Dim i As Long
'        For i = sIdx To eIdx
'            name = name & ChrW$(codes(i))
'        Next i
'
'        ' Если функция — подфункция, рекурсивно вызываем EvalFunction
'        If m_FuncExprCache.Exists(name) Then
'            Set ParseFactor = EvalFunction(name)
'        Else
'            ' Атом
'            Set ParseFactor = CreateAtom(name)
'        End If
'    End If
'End Function


' --- Разделение строки по верхнему уровню скобок ---
Public Function SplitTop(ByVal s As String, ByVal sep As String) As Collection
    Dim res As New Collection, lvl As Long, i As Long, p As Long
    p = 1
    For i = 1 To Len(s)
        Select Case Mid(s, i, 1)
            Case "(": lvl = lvl + 1
            Case ")": lvl = lvl - 1
            Case sep
                If lvl = 0 Then
                    res.Add Mid(s, p, i - p)
                    p = i + 1
                End If
        End Select
    Next i
    res.Add Mid(s, p)
    Set SplitTop = res
End Function


'==============================
' Создание атома (CExpr с одним CTerm)
'==============================


' --- Создание атома ---
Public Function CreateAtom(ByVal sName As String) As CExpr
    Dim res As New CExpr, t As New CTerm, ids() As Long
    ReDim ids(0 To 0)
    ids(0) = GetID(sName)
    t.Init ids, 1, CStr(ids(0))
    res.AddTerm t
    Set CreateAtom = res
End Function

'
'Private Function CreateAtom(ByVal sName As String) As CExpr
'    Dim res As New CExpr
'    Dim t As New CTerm
'    Dim ids(0) As Long
'    ids(0) = GetID(sName)
'    t.Init ids, 1, CStr(ids(0))
'    res.AddTerm t
'    Set CreateAtom = res
'End Function

'==============================
' Объединение выражений (OR)
'==============================
Private Sub MergeExpr(ByRef target As CExpr, ByVal source As CExpr)
    Dim t() As CTerm
    t = source.GetTerms()
    If (Not Not t) = 0 Then Exit Sub
    Dim i As Long
    For i = LBound(t) To UBound(t)
        target.AddTerm t(i)
    Next i
End Sub


'' ===== Разбор функций =====
'Public Function EvalFunction(ByVal fName As String) As CExpr
'    If Not m_FuncExprCache.Exists(fName) Then Err.Raise 998, , "Не найдена функция: " & fName
'    If m_FuncDNFCache.Exists(fName) Then Set EvalFunction = m_FuncDNFCache(fName)
'    If m_CallStack.Exists(fName) Then Err.Raise 997, , "Цикл в функции: " & fName
'
'    m_CallStack.Add fName, True
'    Dim sExpr As String: sExpr = Replace(m_FuncExprCache(fName), " ", "")
'
'    ReDim m_Codes(1 To Len(sExpr))
'    Dim i As Long: For i = 1 To Len(sExpr): m_Codes(i) = AscW(Mid$(sExpr, i, 1)): Next
'
'    Dim res As CExpr: Set res = ParseOr(1, UBound(m_Codes))
'    Set m_FuncDNFCache(fName) = res
'    m_CallStack.Remove fName
'    Set EvalFunction = res
'End Function
'
'Private Function ParseOr(ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    Dim res As New CExpr, i As Long, lvl As Long, lastP As Long
'    lastP = sIdx
'    For i = sIdx To eIdx
'        If m_Codes(i) = AscW(CH_LPAREN) Then lvl = lvl + 1 Else If m_Codes(i) = AscW(CH_RPAREN) Then lvl = lvl - 1
'        If lvl = 0 And m_Codes(i) = AscW(CH_OR) Then
'            MergeExpr res, ParseAnd(lastP, i - 1)
'            lastP = i + 1
'        End If
'    Next
'    MergeExpr res, ParseAnd(lastP, eIdx)
'    Set ParseOr = res
'End Function
'
'Private Function ParseAnd(ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    Dim res As CExpr, parts As New Collection
'    Dim i As Long, lvl As Long, lastP As Long
'    lastP = sIdx
'    For i = sIdx To eIdx
'        If m_Codes(i) = AscW(CH_LPAREN) Then lvl = lvl + 1 Else If m_Codes(i) = AscW(CH_RPAREN) Then lvl = lvl - 1
'        If lvl = 0 And m_Codes(i) = AscW(CH_AND) Then
'            parts.Add Array(lastP, i - 1)
'            lastP = i + 1
'        End If
'    Next
'    parts.Add Array(lastP, eIdx)
'
'    Set res = ParseFactor(CLng(parts(1)(0)), CLng(parts(1)(1)))
'    For i = 2 To parts.Count
'        Set res = MultiplyExpr(res, ParseFactor(CLng(parts(i)(0)), CLng(parts(i)(1))))
'    Next
'    Set ParseAnd = res
'End Function
'

'
'
'Private Function ParseFactor(ByVal sIdx As Long, ByVal eIdx As Long) As CExpr
'    If m_Codes(sIdx) = AscW(CH_LPAREN) And m_Codes(eIdx) = AscW(CH_RPAREN) Then
'        Set ParseFactor = ParseOr(sIdx + 1, eIdx - 1)
'    Else
'        Dim name As String: name = ""
'        Dim i As Long
'        For i = sIdx To eIdx: name = name & ChrW$(m_Codes(i)): Next
'        If m_FuncExprCache.Exists(name) Then
'            Set ParseFactor = EvalFunction(name)
'        Else
'            Set ParseFactor = CreateAtom(name)
'        End If
'    End If
'End Function
'
'Private Function CreateAtom(ByVal sName As String) As CExpr
'    Dim res As New CExpr, t As New CTerm
'    Dim ids() As Long
'    ReDim ids(0 To 0)
'    ids(0) = GetID(sName)
'    t.Init ids, 1, CStr(ids(0))
'    res.AddTerm t
'    Set CreateAtom = res
'End Function
'
'Private Sub MergeExpr(ByRef target As CExpr, ByVal source As CExpr)
'    Dim t() As CTerm: t = source.GetTerms()
'    If (Not Not t) = 0 Then Exit Sub
'    Dim i As Long
'    For i = LBound(t) To UBound(t): target.AddTerm t(i): Next
'End Sub

