Attribute VB_Name = "Writer"
'===========================
' Module: Writer
' Purpose:
'   1) RewriteFailure  : builds symbolic LaTeX from existing DNF engine (EvalFunction -> CExpr -> CTerm[])
'   2) SubstituteFailure: builds numeric LaTeX with Wi and ? substituted from caches
'
' Formatting:
'   All formatting MUST be driven by templates on sheet "Format" (Key in col A, Value in col B).
'   Defaults are provided in code and can be overridden by the sheet.
'
' Dependencies already existing in your project:
'   - InitGlobals
'   - EvalFunction(fName As String) As CExpr
'   - Global caches (CalcFailureMain):
'       Public m_IDToName() As String
'       Public m_LambdaValues() As Double
'       Public m_WiValues() As Double     ' can be 2D via ReDim
'       Public m_FuncExprCache As Object
'       Public m_FuncDNFCache As Object   ' stores CExpr structure
'       Public m_CallStack As Object
'
' Classes already existing:
'   - CExpr with GetTerms() As CTerm()
'   - CTerm with:
'       Multiplier As Double
'       Order As Long
'       Key As String
'       FactorIDs() As Long
'===========================
Option Explicit

'===========================
' Public API
'===========================

' Symbolic LaTeX:
'   Q_{name} = <Wi and lambdas as symbols>
Public Function RewriteFailure(ByVal fName As String, ByVal stage As Variant) As String
    InitGlobals

    Dim expr As CExpr
    Set expr = EvalFunction(fName)

    Dim tpl As Object
    Set tpl = LoadFormatTemplates()

    Dim body As String
    body = RenderExprSymbolicLatex(expr, stage, tpl)

    RewriteFailure = ApplyQNamePrefixLatex(fName, body, tpl)
End Function


' Numeric LaTeX:
'   Q_{name} = <numbers>  where Wi and lambdas are substituted
Public Function SubstituteFailure(ByVal fName As String, ByVal stage As Variant) As String
    InitGlobals

    Dim expr As CExpr
    Set expr = EvalFunction(fName)

    Dim tpl As Object
    Set tpl = LoadFormatTemplates()

    Dim body As String
    body = RenderExprNumericLatex(expr, stage, tpl)

    SubstituteFailure = ApplyQNamePrefixLatex(fName, body, tpl)
End Function


'===========================
' Prefix: Q_{name} = ...
'===========================

Private Function ApplyQNamePrefixLatex(ByVal fName As String, ByVal body As String, ByVal tpl As Object) As String
    Dim prefixTpl As String
    ' Use {FNAME} and {BODY}. Default includes math spacing around "=".
    ' IMPORTANT: keep the whole output valid LaTeX.
    prefixTpl = GetTpl(tpl, "Q_PREFIX_TEMPLATE", "Q_{ {FNAME} }\;=\;{BODY}")

    ApplyQNamePrefixLatex = ApplyTokens(prefixTpl, _
                                        Array("FNAME", "BODY"), _
                                        Array(EscapeLatexText(fName), body))
End Function


'===========================
' Symbolic rendering
'===========================

Private Function RenderExprSymbolicLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim tArr() As CTerm
    tArr = expr.GetTerms()

    If (Not Not tArr) = 0 Then
        RenderExprSymbolicLatex = GetTpl(tpl, "EMPTY_EXPR", "0")
        Exit Function
    End If

    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)

    Dim joinExpr As String
    joinExpr = GetTpl(tpl, "SYM_EXPR_JOIN", " + ")

    Dim out As String
    out = ""

    Dim i As Long
    For i = LBound(tArr) To UBound(tArr)
        Dim part As String
        part = RenderOneCTermSymbolicLatex(tArr(i), stage, tpl)
        If Len(part) > 0 Then
            If Len(out) > 0 Then out = out & joinExpr
            out = out & part
        End If
    Next i

    RenderExprSymbolicLatex = out
End Function

Private Function RenderOneCTermSymbolicLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    ' Пропускаем нулевой множитель (обычно CExpr.AddTerm уже удаляет такие термы)
    If Abs(t.Multiplier) < 0.0000000001 Then
        RenderOneCTermSymbolicLatex = ""
        Exit Function
    End If

    ' --- Шаблоны ---
    ' ВАЖНО: SYM_TERM_TEMPLATE должен содержать {WI_MUL}, например:
    ' {MULT}{WI}{WI_MUL}{LAMPROD}{TP}
    Dim termTpl As String
    termTpl = GetTpl(tpl, "SYM_TERM_TEMPLATE", "{MULT}{WI}{WI_MUL}{LAMPROD}{TP}")

    Dim multTpl As String
    multTpl = GetTpl(tpl, "SYM_MULT_TEMPLATE", "{mult}\,")

    Dim wiTpl As String
    wiTpl = GetTpl(tpl, "SYM_WI_TEMPLATE", "W_{ {r} }^{({stage})}\,")

    Dim lamTpl As String
    lamTpl = GetTpl(tpl, "SYM_LAM_TEMPLATE", "\lambda_{\text{{name}}}")

    Dim lamJoin As String
    lamJoin = GetTpl(tpl, "SYM_LAM_JOIN", "\cdot ")

    ' --- MULT ---
    Dim multStr As String
    multStr = ""
    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumberSymbolic(t.Multiplier)))
    End If

    ' --- WI ---
    Dim wiStr As String
    wiStr = ""
    If Not IsStageAll(stage) Then
        wiStr = ApplyTokens(wiTpl, _
                            Array("r", "stage"), _
                            Array(CStr(t.Order), CStr(stage)))
    End If

    ' --- "умножение после Wi" (только если Wi реально выведен) ---
    Dim wiMulStr As String
    If Len(wiStr) > 0 Then
        wiMulStr = GetTpl(tpl, "SYM_WI_MUL", "\,\cdot\,")
    Else
        wiMulStr = ""
    End If

    ' --- Лямбды ---
    Dim lamProd As String
    lamProd = RenderLambdaProductByIDs(t.FactorIDs, lamTpl, lamJoin)

    ' --- tp (символически): \,t_p  или \,t_p^{r} ---
    Dim tpStr As String
    tpStr = RenderTpSymbolic(t.Order, tpl)

    ' --- Итоговый терм ---
    RenderOneCTermSymbolicLatex = ApplyTokens(termTpl, _
                                              Array("MULT", "WI", "WI_MUL", "LAMPROD", "TP"), _
                                              Array(multStr, wiStr, wiMulStr, lamProd, tpStr))
End Function



'Private Function RenderOneCTermSymbolicLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermSymbolicLatex = ""
'        Exit Function
'    End If
'
'    Dim termTpl As String
'    termTpl = GetTpl(tpl, "SYM_TERM_TEMPLATE", "{MULT}{WI}{LAMPROD}")
'
'    Dim multTpl As String
'    multTpl = GetTpl(tpl, "SYM_MULT_TEMPLATE", "{mult}\,")
'
'    Dim wiTpl As String
'    wiTpl = GetTpl(tpl, "SYM_WI_TEMPLATE", "W_{ {r} }^{({stage})}\,")
'
'    Dim lamTpl As String
'    lamTpl = GetTpl(tpl, "SYM_LAM_TEMPLATE", "\lambda_{\text{{name}}}")
'
'    Dim lamJoin As String
'    lamJoin = GetTpl(tpl, "SYM_LAM_JOIN", "")
'
'    Dim multStr As String
'    multStr = ""
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumberSymbolic(t.Multiplier)))
'    End If
'
'    Dim wiStr As String
'    wiStr = ""
'    If Not IsStageAll(stage) Then
'        wiStr = ApplyTokens(wiTpl, Array("r", "stage"), Array(CStr(t.Order), CStr(stage)))
'    End If
'
'    Dim lamProd As String
'    lamProd = RenderLambdaProductByIDs(t.FactorIDs, lamTpl, lamJoin)
'
'    RenderOneCTermSymbolicLatex = ApplyTokens(termTpl, _
'                                              Array("MULT", "WI", "LAMPROD"), _
'                                              Array(multStr, wiStr, lamProd))
'End Function


'===========================
' Numeric rendering
'===========================

Private Function RenderExprNumericLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim tArr() As CTerm
    tArr = expr.GetTerms()

    If (Not Not tArr) = 0 Then
        RenderExprNumericLatex = GetTpl(tpl, "EMPTY_EXPR", "0")
        Exit Function
    End If

    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)

    Dim joinExpr As String
    joinExpr = GetTpl(tpl, "NUM_EXPR_JOIN", " + ")

    Dim out As String
    out = ""

    Dim i As Long
    For i = LBound(tArr) To UBound(tArr)
        Dim part As String
        part = RenderOneCTermNumericLatex(tArr(i), stage, tpl)
        If Len(part) > 0 Then
            If Len(out) > 0 Then out = out & joinExpr
            out = out & part
        End If
    Next i

    RenderExprNumericLatex = out
End Function

Private Function RenderOneCTermNumericLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    If Abs(t.Multiplier) < 0.0000000001 Then
        RenderOneCTermNumericLatex = ""
        Exit Function
    End If

    ' Как соединять численные множители внутри терма
    Dim factorJoin As String
    factorJoin = GetTpl(tpl, "NUM_FACTOR_JOIN", "\,\cdot\,")

    ' Шаблон терма (ВАЖНО: должен включать {TP}, если вы хотите вывод tp отдельно)
    Dim termTpl As String
    termTpl = GetTpl(tpl, "NUM_TERM_TEMPLATE", "{FACTORS}{TP}")

    Dim factors As Collection
    Set factors = New Collection

    ' Multiplier
    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        factors.Add FormatNumLatex(t.Multiplier, tpl)
    End If

    ' Wi
    Dim wi As Double
    If IsStageAll(stage) Then
        wi = 1#
    Else
        ' В вашем CalcExpr используется m_WiValues(orderIdx, stage)
        ' Здесь придерживаемся того же направления индексов
        If t.Order <= R_MAX Then
            wi = m_WiValues(t.Order, CLng(stage))
        Else
            wi = 0#
        End If
    End If
    If Abs(wi - 1#) > 0.0000000001 Then
        factors.Add FormatNumLatex(wi, tpl)
    End If

    ' Lambdas (только ? — tp добавим отдельным множителем степенью)
    Dim ids() As Long
    ids = t.FactorIDs

    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        factors.Add FormatNumLatex(m_LambdaValues(ids(i)), tpl)
    Next i

    ' Собираем факторы
    Dim factorsStr As String
    factorsStr = JoinCollection(factors, factorJoin)

    ' tp (численно):  \,\cdot\,{tp}  или \,\cdot\,({tp})^{r}
    Dim tpStr As String
    tpStr = RenderTpNumeric(t.Order, tpl)

    RenderOneCTermNumericLatex = ApplyTokens(termTpl, _
                                             Array("FACTORS", "TP"), _
                                             Array(factorsStr, tpStr))
End Function


'Private Function RenderOneCTermNumericLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermNumericLatex = ""
'        Exit Function
'    End If
'
'    Dim factorJoin As String
'    factorJoin = GetTpl(tpl, "NUM_FACTOR_JOIN", "\,\cdot\,")
'
'    Dim termTpl As String
'    termTpl = GetTpl(tpl, "NUM_TERM_TEMPLATE", "{FACTORS}")
'
'    Dim factors As Collection
'    Set factors = New Collection
'
'    ' Multiplier
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(t.Multiplier, tpl)
'    End If
'
'    ' Wi (or 1 if ALL)
'    Dim wi As Double
'    wi = GetWiValue(t.Order, stage)
'    If Abs(wi - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(wi, tpl)
'    End If
'
'    ' Lambdas
'    Dim ids() As Long
'    ids = t.FactorIDs
'
'    Dim i As Long
'    For i = LBound(ids) To UBound(ids)
'        factors.Add FormatNumLatex(GetLambdaValue(ids(i)), tpl)
'    Next i
'
'    Dim factorsStr As String
'    factorsStr = JoinCollection(factors, factorJoin)
'
'    RenderOneCTermNumericLatex = ApplyTokens(termTpl, Array("FACTORS"), Array(factorsStr))
'End Function


'===========================
' Lambda rendering (symbolic)
'===========================

Private Function RenderLambdaProductByIDs(ByRef ids() As Long, ByVal lamTpl As String, ByVal lamJoin As String) As String
    Dim s As String
    s = ""

    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        Dim id As Long: id = ids(i)
        Dim nm As String: nm = GetElementNameByID(id)

        Dim one As String
        one = ApplyTokens(lamTpl, Array("name", "id"), Array(EscapeLatexText(nm), CStr(id)))

        If Len(s) > 0 Then s = s & lamJoin
        s = s & one
    Next i

    RenderLambdaProductByIDs = s
End Function


'===========================
' Numeric getters
'===========================

Private Function GetLambdaValue(ByVal id As Long) As Double
    On Error GoTo EH
    GetLambdaValue = m_LambdaValues(id)
    Exit Function
EH:
    Err.Raise vbObjectError + 701, "Writre.SubstituteFailure", "Нет значения ? для ID=" & CStr(id)
End Function

' stage="ALL" => Wi=1
' Tries 2D access in both orientations: (r,stage) then (stage,r)
Private Function GetWiValue(ByVal r As Long, ByVal stage As Variant) As Double
    If IsStageAll(stage) Then
        GetWiValue = 1#
        Exit Function
    End If

    Dim st As Long
    st = CLng(stage)

    On Error Resume Next

    ' Try (r, st)
    GetWiValue = m_WiValues(r, st)
    If Err.Number = 0 Then Exit Function
    Err.Clear

    ' Try (st, r)
    GetWiValue = m_WiValues(st, r)
    If Err.Number = 0 Then Exit Function
    Err.Clear

    On Error GoTo 0

    Err.Raise vbObjectError + 702, "Writre.SubstituteFailure", _
              "Не удалось прочитать Wi для r=" & CStr(r) & ", stage=" & CStr(st)
End Function


'===========================
' Number formatting (numeric LaTeX)
'  - scientific form as a\cdot 10^{b}
'  - decimal separator: comma (uses locale output from Format$)
'===========================

Private Function FormatNumLatex(ByVal v As Double, ByVal tpl As Object) As String
    Dim plainMin As Double, plainMax As Double
    plainMin = CDblSafe(GetTpl(tpl, "NUM_PLAIN_MIN", "0.001"), 0.001)
    plainMax = CDblSafe(GetTpl(tpl, "NUM_PLAIN_MAX", "1000"), 1000#)

    If v = 0# Then
        FormatNumLatex = "0"
        Exit Function
    End If

    Dim av As Double
    av = Abs(v)

    ' --- plain ---
    If av >= plainMin And av < plainMax Then
        Dim s As String
        s = Format$(v, GetTpl(tpl, "NUM_PLAIN_FMT", "0.############"))

        ' убрать висячую запятую/точку
        If Right$(s, 1) = "," Or Right$(s, 1) = "." Then
            s = Left$(s, Len(s) - 1)
        End If

        FormatNumLatex = s
        Exit Function
    End If

    ' --- scientific ---
    Dim exp As Long
    exp = Fix(Log(av) / Log(10#))

    Dim mant As Double
    mant = v / (10# ^ exp)

    Dim mantFmt As String
    mantFmt = GetTpl(tpl, "NUM_MANTISSA_FMT", "0.#####")

    Dim mantStr As String
    mantStr = Format$(mant, mantFmt)
    If Right$(mantStr, 1) = "," Or Right$(mantStr, 1) = "." Then
        mantStr = Left$(mantStr, Len(mantStr) - 1)
    End If

    Dim sciTpl As String
    sciTpl = GetTpl(tpl, "NUM_SCI_TEMPLATE", "{mant}\cdot 10^{{exp}}")

    FormatNumLatex = ApplyTokens(sciTpl, _
                                 Array("mant", "exp"), _
                                 Array(mantStr, CStr(exp)))
End Function

Private Function CDblSafe(ByVal s As String, ByVal defaultValue As Double) As Double
    On Error GoTo EH
    ' Разрешаем и точку, и запятую в шаблонах
    CDblSafe = CDbl(Replace(s, ".", ","))
    Exit Function
EH:
    CDblSafe = defaultValue
End Function
'===========================
' Templates
'===========================

Private Function LoadFormatTemplates() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' ---------- Common ----------
    d("Q_PREFIX_TEMPLATE") = "Q_{ {FNAME} }\;=\;{BODY}"
    d("EMPTY_EXPR") = "0"

    ' ---------- Symbolic ----------
    d("SYM_EXPR_JOIN") = " + "
    d("SYM_TERM_TEMPLATE") = "{MULT}{WI}{LAMPROD}"
    d("SYM_MULT_TEMPLATE") = "{mult}\,"
    d("SYM_WI_TEMPLATE") = "W_{ {r} }^{({stage})}\,"
    d("SYM_LAM_TEMPLATE") = "\lambda_{\text{{name}}}"
    d("SYM_LAM_JOIN") = ""  ' e.g. "\cdot " or "\," if desired

    ' ---------- Numeric ----------
    d("NUM_EXPR_JOIN") = " + "
    d("NUM_TERM_TEMPLATE") = "{FACTORS}"
    d("NUM_FACTOR_JOIN") = "\,\cdot\,"

    ' number formatting knobs
    d("NUM_PLAIN_MIN") = "0.001"
    d("NUM_PLAIN_MAX") = "1000"
    d("NUM_PLAIN_FMT") = "0.############"
    d("NUM_MANTISSA_FMT") = "0.#####"
    d("NUM_SCI_TEMPLATE") = "{mant}\cdot 10^{{exp}}"

    ' Override from sheet "Format": A=Key, B=Value
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Format")
    On Error GoTo 0

    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim r As Long
        For r = 1 To lastRow
            Dim k As String, v As String
            k = Trim$(CStr(ws.Cells(r, 1).Value))
            v = CStr(ws.Cells(r, 2).Value)
            If Len(k) > 0 Then d(k) = v
        Next r
    End If

    Set LoadFormatTemplates = d
End Function

Private Function GetTpl(ByVal tpl As Object, ByVal key As String, ByVal defaultValue As String) As String
    If Not tpl Is Nothing Then
        If tpl.Exists(key) Then
            GetTpl = CStr(tpl(key))
            Exit Function
        End If
    End If
    GetTpl = defaultValue
End Function

Private Function ApplyTokens(ByVal template As String, ByVal keys As Variant, ByVal values As Variant) As String
    Dim s As String
    s = template

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        s = Replace(s, "{" & CStr(keys(i)) & "}", CStr(values(i)))
    Next i

    ApplyTokens = s
End Function


'===========================
' Sorting (stable output)
'===========================

' Sort by Order ascending, then Key
Private Sub QuickSortCTermArray(ByRef arr() As CTerm, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    i = first: j = last

    Dim pivot As CTerm
    Set pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While CompareCTerm(arr(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareCTerm(arr(j), pivot) > 0
            j = j - 1
        Loop

        If i <= j Then
            Dim tmp As CTerm
            Set tmp = arr(i)
            Set arr(i) = arr(j)
            Set arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortCTermArray arr, first, j
    If i < last Then QuickSortCTermArray arr, i, last
End Sub

Private Function CompareCTerm(ByVal a As CTerm, ByVal b As CTerm) As Long
    If a.Order < b.Order Then
        CompareCTerm = -1
        Exit Function
    ElseIf a.Order > b.Order Then
        CompareCTerm = 1
        Exit Function
    End If
    CompareCTerm = StrComp(a.key, b.key, vbTextCompare)
End Function


'===========================
' Helpers
'===========================

Private Function JoinCollection(ByVal col As Collection, ByVal delim As String) As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        If i > 1 Then s = s & delim
        s = s & CStr(col(i))
    Next i
    JoinCollection = s
End Function

Private Function GetElementNameByID(ByVal id As Long) As String
    On Error GoTo EH
    GetElementNameByID = m_IDToName(id)
    Exit Function
EH:
    GetElementNameByID = "ID" & CStr(id)
End Function

Private Function EscapeLatexText(ByVal x As String) As String
    ' Minimal escaping for \text{...} / indices:
    x = Replace(x, "\", "\\")
    x = Replace(x, "{", "\{")
    x = Replace(x, "}", "\}")
    EscapeLatexText = x
End Function

Private Function IsStageAll(ByVal stage As Variant) As Boolean
    IsStageAll = (VarType(stage) = vbString And UCase$(CStr(stage)) = "ALL")
End Function

Private Function TrimNumberSymbolic(ByVal v As Double) As String
    Dim s As String
    ' Сначала в локальном формате
    s = Format$(v, "0.############")

    ' Нормализуем десятичный разделитель к точке для LaTeX
    s = Replace(s, ",", ".")

    ' Убираем висячую точку, если получилось "2."
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    ' На всякий случай: если вдруг получилось "-0" или "-0."
    If s = "-0" Then s = "0"

    TrimNumberSymbolic = s
End Function

Private Function RenderTpSymbolic(ByVal r As Long, ByVal tpl As Object) As String
    If r <= 0 Then
        RenderTpSymbolic = ""
    ElseIf r = 1 Then
        RenderTpSymbolic = GetTpl(tpl, "TP_SYM_1", "\,t_p")
    Else
        Dim t As String
        t = GetTpl(tpl, "TP_SYM_POW", "\,t_p^{ {r} }")
        RenderTpSymbolic = ApplyTokens(t, Array("r"), Array(CStr(r)))
    End If
End Function

Private Function RenderTpNumeric(ByVal r As Long, ByVal tpl As Object) As String
    If r <= 0 Then
        RenderTpNumeric = ""
        Exit Function
    End If

    Dim tpStr As String
    tpStr = FormatNumLatex(m_Tp, tpl)

    If r = 1 Then
        Dim t1 As String
        t1 = GetTpl(tpl, "TP_NUM_1", "\,\cdot\,{tp}")

        ' Поддержка двух токенов: {tp} и {base}
        RenderTpNumeric = ApplyTokens(t1, _
                                      Array("tp", "base"), _
                                      Array(tpStr, tpStr))
    Else
        Dim powTpl As String
        powTpl = GetTpl(tpl, "TP_NUM_POW", "\,\cdot\,({tp})^{ {r} }")

        ' Поддержка двух токенов: {tp} и {base}, плюс {r}
        RenderTpNumeric = ApplyTokens(powTpl, _
                                      Array("tp", "base", "r"), _
                                      Array(tpStr, tpStr, CStr(r)))
    End If
End Function

''===========================
'' Module: Writer
'' Purpose: Build LaTeX (Word equation friendly) representation of failure formula
''          using existing DNF engine (EvalFunction -> CExpr -> CTerm[])
''
'' Dependencies (already existing in your project):
''   - InitGlobals
''   - EvalFunction(fName As String) As CExpr
''   - Global caches in CalcFailureMain:
''       m_IDToName() As String
''       m_FuncExprCache As Object
''       m_FuncDNFCache As Object
''       m_CallStack As Object
''
'' Data classes (already existing):
''   - Class CExpr with GetTerms() As CTerm()
''   - Class CTerm with:
''         Multiplier As Double
''         Order As Long
''         Key As String
''         FactorIDs() As Long
''
'' Optional:
''   - Worksheet "Format" with Key/Value pairs (A=Key, B=Value)
''===========================
'Option Explicit
'
''---------------------------
'' Public API
''---------------------------
'
'' UDF-friendly: returns a LaTeX string (can be pasted into Word equation)
'Public Function RewriteFailure(ByVal fName As String, ByVal stage As Variant) As String
'    InitGlobals
'
'    Dim expr As CExpr
'    Set expr = EvalFunction(fName)   ' Uses your existing cache/cycle control/parser
'
'    Dim tpl As Object
'    Set tpl = LoadFormatTemplates()
'
'    RewriteFailure = RenderExprToLatex(expr, stage, tpl)
'End Function
'
''---------------------------
'' Rendering: CExpr -> LaTeX
''---------------------------
'
'Private Function RenderExprToLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
'    Dim tArr() As CTerm
'    tArr = expr.GetTerms()
'
'    ' Empty dynamic array check
'    If (Not Not tArr) = 0 Then
'        RenderExprToLatex = ""
'        Exit Function
'    End If
'
'    ' Stable output (dictionary order is not stable)
'    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)
'
'    Dim joinExpr As String
'    joinExpr = GetTpl(tpl, "EXPR_JOIN", " + ")
'
'    Dim out As String
'    out = ""
'
'    Dim i As Long
'    For i = LBound(tArr) To UBound(tArr)
'        Dim part As String
'        part = RenderOneCTermLatex(tArr(i), stage, tpl)
'
'        If Len(part) > 0 Then
'            If Len(out) > 0 Then out = out & joinExpr
'            out = out & part
'        End If
'    Next i
'
'    RenderExprToLatex = out
'End Function
'
'Private Function RenderOneCTermLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
'    ' Safety: skip near-zero multipliers (normally removed by CExpr.AddTerm already)
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermLatex = ""
'        Exit Function
'    End If
'
'    Dim multTpl As String
'    multTpl = GetTpl(tpl, "MULT_TEMPLATE", "{mult}\,")
'
'    Dim wiTpl As String
'    wiTpl = GetTpl(tpl, "WI_TEMPLATE", "W_{{r}}^{{({stage})}}\,") ' default: W_{r}^{(stage)}\,
'
'    Dim termTpl As String
'    termTpl = GetTpl(tpl, "TERM_TEMPLATE", "{MULT}{WI}{LAMPROD}")
'
'    Dim multStr As String
'    multStr = ""
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumber(t.Multiplier)))
'    End If
'
'    Dim wiStr As String
'    wiStr = ""
'    If Not IsStageAll(stage) Then
'        wiStr = ApplyTokens(wiTpl, Array("r", "stage"), Array(CStr(t.Order), CStr(stage)))
'    End If
'
'    Dim lamProd As String
'    lamProd = RenderLambdaProductByIDs(t.FactorIDs, tpl)
'
'    RenderOneCTermLatex = ApplyTokens(termTpl, _
'                                     Array("MULT", "WI", "LAMPROD"), _
'                                     Array(multStr, wiStr, lamProd))
'End Function
'
'Private Function RenderLambdaProductByIDs(ByRef ids() As Long, ByVal tpl As Object) As String
'    Dim lamTpl As String
'    lamTpl = GetTpl(tpl, "LAM_TEMPLATE", "\lambda_{\text{{{name}}}}")
'
'    Dim lamJoin As String
'    lamJoin = GetTpl(tpl, "LAM_JOIN", "") ' default: concatenation
'
'    Dim s As String
'    s = ""
'
'    Dim i As Long
'    For i = LBound(ids) To UBound(ids)
'        Dim id As Long: id = ids(i)
'        Dim nm As String: nm = GetElementNameByID(id)
'
'        Dim one As String
'        one = ApplyTokens(lamTpl, _
'                          Array("name", "id"), _
'                          Array(EscapeLatexText(nm), CStr(id)))
'
'        If Len(s) > 0 Then s = s & lamJoin
'        s = s & one
'    Next i
'
'    RenderLambdaProductByIDs = s
'End Function
'
''---------------------------
'' Templates
''---------------------------
'
'Private Function LoadFormatTemplates() As Object
'    Dim d As Object
'    Set d = CreateObject("Scripting.Dictionary")
'
'    ' Defaults
'    d("EXPR_JOIN") = " + "
'    d("TERM_TEMPLATE") = "{MULT}{WI}{LAMPROD}"
'    d("MULT_TEMPLATE") = "{mult}\,"
'    d("WI_TEMPLATE") = "W_{{r}}^{{({stage})}}\,"
'    d("LAM_TEMPLATE") = "\lambda_{\text{{{name}}}}"
'    d("LAM_JOIN") = ""
'
'    ' Optional override from sheet "Format": A=Key, B=Value
'    On Error Resume Next
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("Format")
'    On Error GoTo 0
'
'    If Not ws Is Nothing Then
'        Dim lastRow As Long
'        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'
'        Dim r As Long
'        For r = 1 To lastRow
'            Dim k As String, v As String
'            k = Trim$(CStr(ws.Cells(r, 1).Value))
'            v = CStr(ws.Cells(r, 2).Value)
'            If Len(k) > 0 Then d(k) = v
'        Next r
'    End If
'
'    Set LoadFormatTemplates = d
'End Function
'
'Private Function GetTpl(ByVal tpl As Object, ByVal key As String, ByVal defaultValue As String) As String
'    If Not tpl Is Nothing Then
'        If tpl.Exists(key) Then
'            GetTpl = CStr(tpl(key))
'            Exit Function
'        End If
'    End If
'    GetTpl = defaultValue
'End Function
'
'Private Function ApplyTokens(ByVal template As String, ByVal keys As Variant, ByVal values As Variant) As String
'    Dim s As String
'    s = template
'
'    Dim i As Long
'    For i = LBound(keys) To UBound(keys)
'        s = Replace(s, "{" & CStr(keys(i)) & "}", CStr(values(i)))
'    Next i
'
'    ApplyTokens = s
'End Function
'
''---------------------------
'' Output stability: sorting
''---------------------------
'
'' Sort by Order (ascending), then by Key (text compare)
'Private Sub QuickSortCTermArray(ByRef arr() As CTerm, ByVal first As Long, ByVal last As Long)
'    Dim i As Long, j As Long
'    i = first: j = last
'
'    Dim pivot As CTerm
'    Set pivot = arr((first + last) \ 2)
'
'    Do While i <= j
'        Do While CompareCTerm(arr(i), pivot) < 0
'            i = i + 1
'        Loop
'
'        Do While CompareCTerm(arr(j), pivot) > 0
'            j = j - 1
'        Loop
'
'        If i <= j Then
'            Dim tmp As CTerm
'            Set tmp = arr(i)
'            Set arr(i) = arr(j)
'            Set arr(j) = tmp
'            i = i + 1
'            j = j - 1
'        End If
'    Loop
'
'    If first < j Then QuickSortCTermArray arr, first, j
'    If i < last Then QuickSortCTermArray arr, i, last
'End Sub
'
'Private Function CompareCTerm(ByVal a As CTerm, ByVal b As CTerm) As Long
'    If a.Order < b.Order Then
'        CompareCTerm = -1
'        Exit Function
'    ElseIf a.Order > b.Order Then
'        CompareCTerm = 1
'        Exit Function
'    End If
'
'    CompareCTerm = StrComp(a.key, b.key, vbTextCompare)
'End Function
'
''---------------------------
'' Helpers
''---------------------------
'
'Private Function GetElementNameByID(ByVal id As Long) As String
'    On Error GoTo EH
'    GetElementNameByID = m_IDToName(id)
'    Exit Function
'EH:
'    GetElementNameByID = "ID" & CStr(id)
'End Function
'
'Private Function EscapeLatexText(ByVal x As String) As String
'    ' Minimal escaping for \text{...}
'    x = Replace(x, "\", "\\")
'    x = Replace(x, "{", "\{")
'    x = Replace(x, "}", "\}")
'    EscapeLatexText = x
'End Function
'
'Private Function IsStageAll(ByVal stage As Variant) As Boolean
'    IsStageAll = (VarType(stage) = vbString And UCase$(CStr(stage)) = "ALL")
'End Function
'
'Private Function TrimNumber(ByVal v As Double) As String
'    ' Remove trailing zeros, enforce dot as decimal separator
'    TrimNumber = Replace(Format$(v, "0.############"), ",", ".")
'End Function
'
'Option Explicit
'
''===========================
'' Public: SubstituteFailure
''===========================
'' Подставляет числовые значения Wi и ? в формулу отказа (LaTeX строка)
'' Пример вывода:
''   0.85\,\cdot\,1.2E-6\,\cdot\,3.4E-7 + ...
''
'Public Function SubstituteFailure(ByVal fName As String, ByVal stage As Variant) As String
'    InitGlobals
'
'    Dim expr As CExpr
'    Set expr = EvalFunction(fName)
'
'    Dim tArr() As CTerm
'    tArr = expr.GetTerms()
'    If (Not Not tArr) = 0 Then
'        SubstituteFailure = ""
'        Exit Function
'    End If
'
'    ' Чтобы порядок был стабильным
'    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)
'
'    Dim out As String: out = ""
'    Dim joinExpr As String: joinExpr = " + "   ' можно также вынести в Format как EXPR_JOIN_NUM
'
'    Dim i As Long
'    For i = LBound(tArr) To UBound(tArr)
'        Dim part As String
'        part = RenderOneCTermNumericLatex(tArr(i), stage)
'        If Len(part) > 0 Then
'            If Len(out) > 0 Then out = out & joinExpr
'            out = out & part
'        End If
'    Next i
'
'    SubstituteFailure = out
'End Function
'
'
''===========================
'' Term -> numeric LaTeX
''===========================
'Private Function RenderOneCTermNumericLatex(ByVal t As CTerm, ByVal stage As Variant) As String
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermNumericLatex = ""
'        Exit Function
'    End If
'
'    Dim factors As Collection
'    Set factors = New Collection
'
'    ' Multiplier
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(t.Multiplier)
'    End If
'
'    ' Wi (или 1 в ALL)
'    Dim wi As Double
'    wi = GetWiValue(t.Order, stage)
'    If Abs(wi - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(wi)
'    End If
'
'    ' Lambdas
'    Dim ids() As Long
'    ids = t.FactorIDs
'
'    Dim i As Long
'    For i = LBound(ids) To UBound(ids)
'        Dim lam As Double
'        lam = GetLambdaValue(ids(i))
'        factors.Add FormatNumLatex(lam)
'    Next i
'
'    RenderOneCTermNumericLatex = JoinCollectionLatex(factors, "\,\cdot\,")
'End Function
'
'
''===========================
'' Numeric getters
''===========================
'
'Private Function GetLambdaValue(ByVal id As Long) As Double
'    On Error GoTo EH
'    GetLambdaValue = m_LambdaValues(id)
'    Exit Function
'EH:
'    Err.Raise vbObjectError + 701, "SubstituteFailure", "Нет ? для ID=" & id
'End Function
'
'
'' Wi: stage="ALL" => 1
'' Очень “устойчивое” чтение m_WiValues:
''  - пробуем m_WiValues(r, stage)
''  - пробуем m_WiValues(stage, r)
''  - пробуем 1D (плоский) как stage-major или r-major
'Private Function GetWiValue(ByVal r As Long, ByVal stage As Variant) As Double
'    If IsStageAll(stage) Then
'        GetWiValue = 1#
'        Exit Function
'    End If
'
'    Dim st As Long
'    st = CLng(stage)
'
'    ' 1) 2D: (r, stage)
'    On Error Resume Next
'    GetWiValue = m_WiValues(r, st)
'    If Err.Number = 0 Then Exit Function
'    Err.Clear
'
'    ' 2) 2D: (stage, r)
'    GetWiValue = m_WiValues(st, r)
'    If Err.Number = 0 Then Exit Function
'    Err.Clear
'    On Error GoTo 0
'
'    ' 3) 1D fallback (если Wi хранится плоско)
'    ' Попробуем угадать раскладку. Обычно r = 0..R_MAX, stage=0..12
'    ' Здесь используем известные границы, если они есть.
'    Dim rMax As Long: rMax = GuessRMax()
'    Dim sMax As Long: sMax = GuessStageMax()
'
'    ' stage-major: idx = st*(rMax+1) + r
'    Dim idx As Long
'    idx = st * (rMax + 1) + r
'    If CanReadWi1D(idx) Then
'        GetWiValue = m_WiValues(idx)
'        Exit Function
'    End If
'
'    ' r-major: idx = r*(sMax+1) + st
'    idx = r * (sMax + 1) + st
'    If CanReadWi1D(idx) Then
'        GetWiValue = m_WiValues(idx)
'        Exit Function
'    End If
'
'    Err.Raise vbObjectError + 702, "SubstituteFailure", _
'              "Не удалось прочитать Wi для r=" & r & ", stage=" & st & " из m_WiValues"
'End Function
'
'Private Function CanReadWi1D(ByVal idx As Long) As Boolean
'    On Error Resume Next
'    Dim v As Double
'    v = m_WiValues(idx)
'    CanReadWi1D = (Err.Number = 0)
'    Err.Clear
'    On Error GoTo 0
'End Function
'
'Private Function GuessRMax() As Long
'    ' Если у тебя есть константа R_MAX — используем её.
'    On Error GoTo fallback
'    GuessRMax = CLng(Evaluate("R_MAX"))
'    Exit Function
'fallback:
'    ' безопасный дефолт
'    GuessRMax = 12
'End Function
'
'Private Function GuessStageMax() As Long
'    ' Обычно 0..12
'    GuessStageMax = 12
'End Function
'
'
''===========================
'' Formatting helpers
''===========================
'
'' Формат числа для LaTeX: научный вид для маленьких/больших значений
'' Word обычно нормально понимает "1.2E-6" в формуле
'Private Function FormatNumLatex(ByVal v As Double) As String
'    If v = 0# Then
'        FormatNumLatex = "0"
'        Exit Function
'    End If
'
'    Dim av As Double
'    av = Abs(v)
'
'    ' Обычная запись — без степени
'    If av >= 0.001 And av < 1000# Then
'        Dim s As String
'        s = Format$(v, "0.############")
'        ' НИЧЕГО не заменяем: Excel сам использует запятую
'        FormatNumLatex = s
'        Exit Function
'    End If
'
'    ' Научная запись: a·10^{b}
'    Dim exp As Long
'    exp = Fix(Log(av) / Log(10#))
'
'    Dim mant As Double
'    mant = v / (10# ^ exp)
'
'    Dim mantStr As String
'    mantStr = Format$(mant, "0.#####")
'    ' запятая остаётся запятой
'
'    FormatNumLatex = mantStr & "\cdot 10^{" & CStr(exp) & "}"
'End Function
'
'
'Private Function JoinCollectionLatex(ByVal col As Collection, ByVal delim As String) As String
'    Dim i As Long, s As String
'    For i = 1 To col.Count
'        If i > 1 Then s = s & delim
'        s = s & CStr(col(i))
'    Next
'    JoinCollectionLatex = s
'End Function
'

''===========================
'' Module: Writer
'' Purpose: Build LaTeX (Word equation friendly) representation of failure formula
''          using existing DNF engine (EvalFunction -> CExpr -> CTerm[])
''
'' Dependencies (already existing in your project):
''   - InitGlobals
''   - EvalFunction(fName As String) As CExpr
''   - Global caches in CalcFailureMain:
''       m_IDToName() As String
''       m_FuncExprCache As Object
''       m_FuncDNFCache As Object
''       m_CallStack As Object
''
'' Data classes (already existing):
''   - Class CExpr with GetTerms() As CTerm()
''   - Class CTerm with:
''         Multiplier As Double
''         Order As Long
''         Key As String
''         FactorIDs() As Long
''
'' Optional:
''   - Worksheet "Format" with Key/Value pairs (A=Key, B=Value)
''===========================
'Option Explicit
'
''---------------------------
'' Public API
''---------------------------
'
'' UDF-friendly: returns a LaTeX string (can be pasted into Word equation)
'Public Function RewriteFailure(ByVal fName As String, ByVal stage As Variant) As String
'    InitGlobals
'
'    Dim expr As CExpr
'    Set expr = EvalFunction(fName)   ' Uses your existing cache/cycle control/parser
'
'    Dim tpl As Object
'    Set tpl = LoadFormatTemplates()
'
'    RewriteFailure = RenderExprToLatex(expr, stage, tpl)
'End Function
'
''---------------------------
'' Rendering: CExpr -> LaTeX
''---------------------------
'
'Private Function RenderExprToLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
'    Dim tArr() As CTerm
'    tArr = expr.GetTerms()
'
'    ' Empty dynamic array check
'    If (Not Not tArr) = 0 Then
'        RenderExprToLatex = ""
'        Exit Function
'    End If
'
'    ' Stable output (dictionary order is not stable)
'    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)
'
'    Dim joinExpr As String
'    joinExpr = GetTpl(tpl, "EXPR_JOIN", " + ")
'
'    Dim out As String
'    out = ""
'
'    Dim i As Long
'    For i = LBound(tArr) To UBound(tArr)
'        Dim part As String
'        part = RenderOneCTermLatex(tArr(i), stage, tpl)
'
'        If Len(part) > 0 Then
'            If Len(out) > 0 Then out = out & joinExpr
'            out = out & part
'        End If
'    Next i
'
'    RenderExprToLatex = out
'End Function
'
'Private Function RenderOneCTermLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
'    ' Safety: skip near-zero multipliers (normally removed by CExpr.AddTerm already)
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermLatex = ""
'        Exit Function
'    End If
'
'    Dim multTpl As String
'    multTpl = GetTpl(tpl, "MULT_TEMPLATE", "{mult}\,")
'
'    Dim wiTpl As String
'    wiTpl = GetTpl(tpl, "WI_TEMPLATE", "W_{{r}}^{{({stage})}}\,") ' default: W_{r}^{(stage)}\,
'
'    Dim termTpl As String
'    termTpl = GetTpl(tpl, "TERM_TEMPLATE", "{MULT}{WI}{LAMPROD}")
'
'    Dim multStr As String
'    multStr = ""
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumber(t.Multiplier)))
'    End If
'
'    Dim wiStr As String
'    wiStr = ""
'    If Not IsStageAll(stage) Then
'        wiStr = ApplyTokens(wiTpl, Array("r", "stage"), Array(CStr(t.Order), CStr(stage)))
'    End If
'
'    Dim lamProd As String
'    lamProd = RenderLambdaProductByIDs(t.FactorIDs, tpl)
'
'    RenderOneCTermLatex = ApplyTokens(termTpl, _
'                                     Array("MULT", "WI", "LAMPROD"), _
'                                     Array(multStr, wiStr, lamProd))
'End Function
'
'Private Function RenderLambdaProductByIDs(ByRef ids() As Long, ByVal tpl As Object) As String
'    Dim lamTpl As String
'    lamTpl = GetTpl(tpl, "LAM_TEMPLATE", "\lambda_{\text{{{name}}}}")
'
'    Dim lamJoin As String
'    lamJoin = GetTpl(tpl, "LAM_JOIN", "") ' default: concatenation
'
'    Dim s As String
'    s = ""
'
'    Dim i As Long
'    For i = LBound(ids) To UBound(ids)
'        Dim id As Long: id = ids(i)
'        Dim nm As String: nm = GetElementNameByID(id)
'
'        Dim one As String
'        one = ApplyTokens(lamTpl, _
'                          Array("name", "id"), _
'                          Array(EscapeLatexText(nm), CStr(id)))
'
'        If Len(s) > 0 Then s = s & lamJoin
'        s = s & one
'    Next i
'
'    RenderLambdaProductByIDs = s
'End Function
'
''---------------------------
'' Templates
''---------------------------
'
'Private Function LoadFormatTemplates() As Object
'    Dim d As Object
'    Set d = CreateObject("Scripting.Dictionary")
'
'    ' Defaults
'    d("EXPR_JOIN") = " + "
'    d("TERM_TEMPLATE") = "{MULT}{WI}{LAMPROD}"
'    d("MULT_TEMPLATE") = "{mult}\,"
'    d("WI_TEMPLATE") = "W_{{r}}^{{({stage})}}\,"
'    d("LAM_TEMPLATE") = "\lambda_{\text{{{name}}}}"
'    d("LAM_JOIN") = ""
'
'    ' Optional override from sheet "Format": A=Key, B=Value
'    On Error Resume Next
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("Format")
'    On Error GoTo 0
'
'    If Not ws Is Nothing Then
'        Dim lastRow As Long
'        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'
'        Dim r As Long
'        For r = 1 To lastRow
'            Dim k As String, v As String
'            k = Trim$(CStr(ws.Cells(r, 1).Value))
'            v = CStr(ws.Cells(r, 2).Value)
'            If Len(k) > 0 Then d(k) = v
'        Next r
'    End If
'
'    Set LoadFormatTemplates = d
'End Function
'
'Private Function GetTpl(ByVal tpl As Object, ByVal key As String, ByVal defaultValue As String) As String
'    If Not tpl Is Nothing Then
'        If tpl.Exists(key) Then
'            GetTpl = CStr(tpl(key))
'            Exit Function
'        End If
'    End If
'    GetTpl = defaultValue
'End Function
'
'Private Function ApplyTokens(ByVal template As String, ByVal keys As Variant, ByVal values As Variant) As String
'    Dim s As String
'    s = template
'
'    Dim i As Long
'    For i = LBound(keys) To UBound(keys)
'        s = Replace(s, "{" & CStr(keys(i)) & "}", CStr(values(i)))
'    Next i
'
'    ApplyTokens = s
'End Function
'
''---------------------------
'' Output stability: sorting
''---------------------------
'
'' Sort by Order (ascending), then by Key (text compare)
'Private Sub QuickSortCTermArray(ByRef arr() As CTerm, ByVal first As Long, ByVal last As Long)
'    Dim i As Long, j As Long
'    i = first: j = last
'
'    Dim pivot As CTerm
'    Set pivot = arr((first + last) \ 2)
'
'    Do While i <= j
'        Do While CompareCTerm(arr(i), pivot) < 0
'            i = i + 1
'        Loop
'
'        Do While CompareCTerm(arr(j), pivot) > 0
'            j = j - 1
'        Loop
'
'        If i <= j Then
'            Dim tmp As CTerm
'            Set tmp = arr(i)
'            Set arr(i) = arr(j)
'            Set arr(j) = tmp
'            i = i + 1
'            j = j - 1
'        End If
'    Loop
'
'    If first < j Then QuickSortCTermArray arr, first, j
'    If i < last Then QuickSortCTermArray arr, i, last
'End Sub
'
'Private Function CompareCTerm(ByVal a As CTerm, ByVal b As CTerm) As Long
'    If a.Order < b.Order Then
'        CompareCTerm = -1
'        Exit Function
'    ElseIf a.Order > b.Order Then
'        CompareCTerm = 1
'        Exit Function
'    End If
'
'    CompareCTerm = StrComp(a.key, b.key, vbTextCompare)
'End Function
'
''---------------------------
'' Helpers
''---------------------------
'
'Private Function GetElementNameByID(ByVal id As Long) As String
'    On Error GoTo EH
'    GetElementNameByID = m_IDToName(id)
'    Exit Function
'EH:
'    GetElementNameByID = "ID" & CStr(id)
'End Function
'
'Private Function EscapeLatexText(ByVal x As String) As String
'    ' Minimal escaping for \text{...}
'    x = Replace(x, "\", "\\")
'    x = Replace(x, "{", "\{")
'    x = Replace(x, "}", "\}")
'    EscapeLatexText = x
'End Function
'
'Private Function IsStageAll(ByVal stage As Variant) As Boolean
'    IsStageAll = (VarType(stage) = vbString And UCase$(CStr(stage)) = "ALL")
'End Function
'
'Private Function TrimNumber(ByVal v As Double) As String
'    ' Remove trailing zeros, enforce dot as decimal separator
'    TrimNumber = Replace(Format$(v, "0.############"), ",", ".")
'End Function
'
'Option Explicit
'
''===========================
'' Public: SubstituteFailure
''===========================
'' Подставляет числовые значения Wi и ? в формулу отказа (LaTeX строка)
'' Пример вывода:
''   0.85\,\cdot\,1.2E-6\,\cdot\,3.4E-7 + ...
''
'Public Function SubstituteFailure(ByVal fName As String, ByVal stage As Variant) As String
'    InitGlobals
'
'    Dim expr As CExpr
'    Set expr = EvalFunction(fName)
'
'    Dim tArr() As CTerm
'    tArr = expr.GetTerms()
'    If (Not Not tArr) = 0 Then
'        SubstituteFailure = ""
'        Exit Function
'    End If
'
'    ' Чтобы порядок был стабильным
'    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)
'
'    Dim out As String: out = ""
'    Dim joinExpr As String: joinExpr = " + "   ' можно также вынести в Format как EXPR_JOIN_NUM
'
'    Dim i As Long
'    For i = LBound(tArr) To UBound(tArr)
'        Dim part As String
'        part = RenderOneCTermNumericLatex(tArr(i), stage)
'        If Len(part) > 0 Then
'            If Len(out) > 0 Then out = out & joinExpr
'            out = out & part
'        End If
'    Next i
'
'    SubstituteFailure = out
'End Function
'
'
''===========================
'' Term -> numeric LaTeX
''===========================
'Private Function RenderOneCTermNumericLatex(ByVal t As CTerm, ByVal stage As Variant) As String
'    If Abs(t.Multiplier) < 0.0000000001 Then
'        RenderOneCTermNumericLatex = ""
'        Exit Function
'    End If
'
'    Dim factors As Collection
'    Set factors = New Collection
'
'    ' Multiplier
'    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(t.Multiplier)
'    End If
'
'    ' Wi (или 1 в ALL)
'    Dim wi As Double
'    wi = GetWiValue(t.Order, stage)
'    If Abs(wi - 1#) > 0.0000000001 Then
'        factors.Add FormatNumLatex(wi)
'    End If
'
'    ' Lambdas
'    Dim ids() As Long
'    ids = t.FactorIDs
'
'    Dim i As Long
'    For i = LBound(ids) To UBound(ids)
'        Dim lam As Double
'        lam = GetLambdaValue(ids(i))
'        factors.Add FormatNumLatex(lam)
'    Next i
'
'    RenderOneCTermNumericLatex = JoinCollectionLatex(factors, "\,\cdot\,")
'End Function
'
'
''===========================
'' Numeric getters
''===========================
'
'Private Function GetLambdaValue(ByVal id As Long) As Double
'    On Error GoTo EH
'    GetLambdaValue = m_LambdaValues(id)
'    Exit Function
'EH:
'    Err.Raise vbObjectError + 701, "SubstituteFailure", "Нет ? для ID=" & id
'End Function
'
'
'' Wi: stage="ALL" => 1
'' Очень “устойчивое” чтение m_WiValues:
''  - пробуем m_WiValues(r, stage)
''  - пробуем m_WiValues(stage, r)
''  - пробуем 1D (плоский) как stage-major или r-major
'Private Function GetWiValue(ByVal r As Long, ByVal stage As Variant) As Double
'    If IsStageAll(stage) Then
'        GetWiValue = 1#
'        Exit Function
'    End If
'
'    Dim st As Long
'    st = CLng(stage)
'
'    ' 1) 2D: (r, stage)
'    On Error Resume Next
'    GetWiValue = m_WiValues(r, st)
'    If Err.Number = 0 Then Exit Function
'    Err.Clear
'
'    ' 2) 2D: (stage, r)
'    GetWiValue = m_WiValues(st, r)
'    If Err.Number = 0 Then Exit Function
'    Err.Clear
'    On Error GoTo 0
'
'    ' 3) 1D fallback (если Wi хранится плоско)
'    ' Попробуем угадать раскладку. Обычно r = 0..R_MAX, stage=0..12
'    ' Здесь используем известные границы, если они есть.
'    Dim rMax As Long: rMax = GuessRMax()
'    Dim sMax As Long: sMax = GuessStageMax()
'
'    ' stage-major: idx = st*(rMax+1) + r
'    Dim idx As Long
'    idx = st * (rMax + 1) + r
'    If CanReadWi1D(idx) Then
'        GetWiValue = m_WiValues(idx)
'        Exit Function
'    End If
'
'    ' r-major: idx = r*(sMax+1) + st
'    idx = r * (sMax + 1) + st
'    If CanReadWi1D(idx) Then
'        GetWiValue = m_WiValues(idx)
'        Exit Function
'    End If
'
'    Err.Raise vbObjectError + 702, "SubstituteFailure", _
'              "Не удалось прочитать Wi для r=" & r & ", stage=" & st & " из m_WiValues"
'End Function
'
'Private Function CanReadWi1D(ByVal idx As Long) As Boolean
'    On Error Resume Next
'    Dim v As Double
'    v = m_WiValues(idx)
'    CanReadWi1D = (Err.Number = 0)
'    Err.Clear
'    On Error GoTo 0
'End Function
'
'Private Function GuessRMax() As Long
'    ' Если у тебя есть константа R_MAX — используем её.
'    On Error GoTo fallback
'    GuessRMax = CLng(Evaluate("R_MAX"))
'    Exit Function
'fallback:
'    ' безопасный дефолт
'    GuessRMax = 12
'End Function
'
'Private Function GuessStageMax() As Long
'    ' Обычно 0..12
'    GuessStageMax = 12
'End Function
'
'
''===========================
'' Formatting helpers
''===========================
'
'' Формат числа для LaTeX: научный вид для маленьких/больших значений
'' Word обычно нормально понимает "1.2E-6" в формуле
'Private Function FormatNumLatex(ByVal v As Double) As String
'    If v = 0# Then
'        FormatNumLatex = "0"
'        Exit Function
'    End If
'
'    Dim av As Double
'    av = Abs(v)
'
'    ' Обычная запись — без степени
'    If av >= 0.001 And av < 1000# Then
'        Dim s As String
'        s = Format$(v, "0.############")
'        ' НИЧЕГО не заменяем: Excel сам использует запятую
'        FormatNumLatex = s
'        Exit Function
'    End If
'
'    ' Научная запись: a·10^{b}
'    Dim exp As Long
'    exp = Fix(Log(av) / Log(10#))
'
'    Dim mant As Double
'    mant = v / (10# ^ exp)
'
'    Dim mantStr As String
'    mantStr = Format$(mant, "0.#####")
'    ' запятая остаётся запятой
'
'    FormatNumLatex = mantStr & "\cdot 10^{" & CStr(exp) & "}"
'End Function
'
'
'Private Function JoinCollectionLatex(ByVal col As Collection, ByVal delim As String) As String
'    Dim i As Long, s As String
'    For i = 1 To col.Count
'        If i > 1 Then s = s & delim
'        s = s & CStr(col(i))
'    Next
'    JoinCollectionLatex = s
'End Function
'
