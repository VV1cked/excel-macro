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



