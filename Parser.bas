Attribute VB_Name = "Parser"
Option Explicit

'==============================
' Модуль: Parser + EvalFunction (DEBUG VERSION)
'==============================

' Символы
Private Const CH_LPAREN As String = "("
Private Const CH_RPAREN As String = ")"
Private Const CH_PLUS As String = "+"
Private Const CH_MULT As String = "*"

' Коды ошибок
Private Const ERR_SYNTAX As Long = vbObjectError + 1001
Private Const ERR_FUNC_NOT_FOUND As Long = vbObjectError + 3001
Private Const ERR_NAME_NOT_FOUND As Long = vbObjectError + 3002
Private Const ERR_Q_NOT_FOUND As Long = vbObjectError + 3003
Private Const ERR_CYCLE As Long = vbObjectError + 3004

' DEBUG: Глобальная переменная для трассировки
Public g_DebugTrace As Boolean

'=========================================================
' Вычисление функции с учётом кэша и рекурсивных вызовов
'=========================================================

Public Function EvalFunction(ByVal fName As String) As CExpr
    If Not m_FuncExprCache.Exists(fName) Then Err.Raise ERR_FUNC_NOT_FOUND, "EvalFunction", "Не найдена функция: " & fName
    If m_FuncDNFCache.Exists(fName) Then Set EvalFunction = m_FuncDNFCache(fName): Exit Function
    If m_CallStack.Exists(fName) Then Err.Raise ERR_CYCLE, "EvalFunction", "Цикл: " & fName

    m_CallStack.Add fName, True
    Dim sExpr As String: sExpr = Replace(m_FuncExprCache(fName), " ", "")

    On Error GoTo EH
    Dim res As CExpr
    Set res = ParseOr(sExpr, fName)
    Set m_FuncDNFCache(fName) = res
    m_CallStack.Remove fName
    Set EvalFunction = res
    Exit Function

EH:
    Dim msg As String
    msg = "Функция: " & fName & vbCrLf & Err.Description
    m_CallStack.Remove fName
    Err.Raise Err.Number, Err.Source, msg
End Function

'==============================
' Парсер OR выражения
'==============================

Public Function ParseOr(ByVal s As String, ByVal ctxName As String) As CExpr
    If g_DebugTrace Then Debug.Print "ParseOr: """ & s & """ (ctx: " & ctxName & ")"
    
    Dim res As New CExpr, parts As Collection, p As Variant
    Set parts = SplitTop(s, "+", ctxName)
    
    If g_DebugTrace Then Debug.Print "  SplitTop(+) вернул " & parts.Count & " частей"
    
    For Each p In parts
        If g_DebugTrace Then Debug.Print "    Часть: """ & CStr(p) & """"
        Set res = OrExpr(res, ParseAnd(CStr(p), ctxName))
    Next p
    Set ParseOr = res
End Function

'==============================
' Парсер AND выражения
'==============================

Public Function ParseAnd(ByVal s As String, ByVal ctxName As String) As CExpr
    If g_DebugTrace Then Debug.Print "ParseAnd: """ & s & """ (ctx: " & ctxName & ")"
    
    Dim res As New CExpr, parts As Collection, p As Variant
    Set parts = SplitTop(s, "*", ctxName)
    
    If g_DebugTrace Then Debug.Print "  SplitTop(*) вернул " & parts.Count & " частей"
    
    Dim first As Boolean: first = True
    For Each p In parts
        If g_DebugTrace Then Debug.Print "    Часть: """ & CStr(p) & """"
        If first Then
            Set res = ParseFactor(CStr(p), ctxName)
            first = False
        Else
            Set res = MultiplyExpr(res, ParseFactor(CStr(p), ctxName))
        End If
    Next p
    Set ParseAnd = res
End Function

'==============================
' Парсер фактора (атом или скобки)
'==============================

Public Function ParseFactor(ByVal s As String, ByVal ctxName As String) As CExpr
    If g_DebugTrace Then Debug.Print "ParseFactor: """ & s & """ (ctx: " & ctxName & ")"
    
    If Len(s) = 0 Then
        Err.Raise ERR_SYNTAX, "Parser", "Пустой фактор в выражении." & vbCrLf & "Функция: " & ctxName
    End If

    If Left$(s, 1) = "(" Then
        ' Проверяем outer-parens
        If Not IsOuterParens(s) Then
            Err.Raise ERR_SYNTAX, "Parser", "Некорректные скобки в факторе: " & s & vbCrLf & "Функция: " & ctxName
        End If
        
        If g_DebugTrace Then Debug.Print "  Outer parens detected, recursing into ParseOr"
        Set ParseFactor = ParseOr(Mid$(s, 2, Len(s) - 2), ctxName)
        Exit Function
    End If

    ' ИСПРАВЛЕНИЕ: Проверка на лишние скобки должна быть ПОСЛЕ проверки outer-parens
    ' Иначе выражения вида "name(" или "name)" будут ошибочно отклонены
    If InStr(1, s, "(", vbBinaryCompare) > 0 Or InStr(1, s, ")", vbBinaryCompare) > 0 Then
        Err.Raise ERR_SYNTAX, "Parser", "Лишняя скобка в атоме: " & s & vbCrLf & "Функция: " & ctxName
    End If

    ' Вызов функции
    If m_FuncExprCache.Exists(s) Then
        If g_DebugTrace Then Debug.Print "  Detected function reference: " & s
        Set ParseFactor = EvalFunction(s)
        Exit Function
    End If

    ' Атом
    If g_DebugTrace Then Debug.Print "  Creating atom: " & s
    Set ParseFactor = CreateAtomStrict(s, ctxName)
End Function

' --- Разделение строки по верхнему уровню скобок ---
Public Function SplitTop(ByVal s As String, ByVal sep As String, ByVal ctxName As String) As Collection
    If g_DebugTrace Then Debug.Print "SplitTop: """ & s & """, sep='" & sep & "'"
    
    Dim res As New Collection
    Dim lvl As Long, i As Long, p As Long
    p = 1
    lvl = 0  ' ИСПРАВЛЕНИЕ: явная инициализация

    If Len(s) = 0 Then
        Err.Raise ERR_SYNTAX, "Parser", "Пустое выражение." & vbCrLf & "Функция: " & ctxName
    End If

    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        
        If g_DebugTrace And (ch = "(" Or ch = ")" Or ch = sep) Then
            Debug.Print "  i=" & i & ", ch='" & ch & "', lvl=" & lvl & ", p=" & p
        End If
        
        Select Case ch
            Case "("
                lvl = lvl + 1
            Case ")"
                lvl = lvl - 1
                If lvl < 0 Then
                    Err.Raise ERR_SYNTAX, "Parser", _
                        "Лишняя закрывающая скобка ) (позиция " & i & ")." & vbCrLf & MarkPos(s, i) & vbCrLf & "Функция: " & ctxName
                End If
            Case sep
                If lvl = 0 Then
                    Dim part As String
                    part = Mid$(s, p, i - p)
                    
                    If g_DebugTrace Then Debug.Print "  Найден разделитель на lvl=0, извлекаем часть: """ & part & """"
                    
                    If Len(part) = 0 Then
                        Err.Raise ERR_SYNTAX, "Parser", _
                            "Оператор '" & sep & "' без операнда (позиция " & i & ")." & vbCrLf & MarkPos(s, i) & vbCrLf & "Функция: " & ctxName
                    End If
                    res.Add part
                    p = i + 1
                End If
        End Select
    Next i

    If lvl <> 0 Then
        Err.Raise ERR_SYNTAX, "Parser", _
            "Не закрыта скобка ) в выражении (lvl=" & lvl & ")." & vbCrLf & MarkPos(s, Len(s)) & vbCrLf & "Функция: " & ctxName
    End If

    Dim lastPart As String
    lastPart = Mid$(s, p)
    
    If g_DebugTrace Then Debug.Print "  Последняя часть (p=" & p & "): """ & lastPart & """"
    
    If Len(lastPart) = 0 Then
        Err.Raise ERR_SYNTAX, "Parser", _
            "Оператор '" & sep & "' в конце выражения." & vbCrLf & MarkPos(s, Len(s)) & vbCrLf & "Функция: " & ctxName
    End If
    res.Add lastPart

    If g_DebugTrace Then Debug.Print "  SplitTop результат: " & res.Count & " частей"
    
    Set SplitTop = res
End Function

Private Function IsOuterParens(ByVal s As String) As Boolean
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "(" Or Right$(s, 1) <> ")" Then Exit Function

    Dim lvl As Long, i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = "(" Then lvl = lvl + 1
        If ch = ")" Then
            lvl = lvl - 1
            ' Если закрылись в ноль до конца строки — значит внешние скобки не охватывают всё
            If lvl = 0 And i < Len(s) Then Exit Function
        End If
    Next i
    IsOuterParens = (lvl = 0)
End Function

'==============================
' Создание атома (CExpr с одним CTerm)
'==============================

Public Function CreateAtomStrict(ByVal sName As String, Optional ByVal ctx As String = "") As CExpr
    Dim nm As String
    nm = Trim$(sName)

    If Len(nm) = 0 Then
        Err.Raise vbObjectError + 1001, "Parser", "Пустой атом. " & ctx
    End If

    ' Быстрая защита от мусора в атоме (операторы/скобки)
    If InStr(1, nm, "+", vbBinaryCompare) > 0 Or _
       InStr(1, nm, "*", vbBinaryCompare) > 0 Or _
       InStr(1, nm, "(", vbBinaryCompare) > 0 Or _
       InStr(1, nm, ")", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1001, "Parser", "Недопустимые символы в атоме: '" & nm & "'. " & ctx
    End If

    ' Если это имя функции — атомом его не считаем, а считаем ссылкой на функцию
    If Not m_FuncExprCache Is Nothing Then
        If m_FuncExprCache.Exists(nm) Then
            Set CreateAtomStrict = EvalFunction(nm)
            Exit Function
        End If
    End If

    ' Иначе это должен быть элемент или внешняя Q, уже зарегистрированные на этапе загрузки
    Dim id As Long
    id = GetIDStrict(nm, ctx)

    Dim res As New CExpr
    Dim t As New CTerm
    Dim ids() As Long
    ReDim ids(0 To 0)
    ids(0) = id

    t.Init ids, 1, CStr(id)
    res.AddTerm t

    Set CreateAtomStrict = res
End Function

Private Function MarkPos(ByVal s As String, ByVal pos As Long) As String
    If pos < 1 Then pos = 1
    If pos > Len(s) Then pos = Len(s)
    MarkPos = s & vbCrLf & Space$(pos - 1) & "^"
End Function
