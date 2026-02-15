Attribute VB_Name = "Writer"
Option Explicit

Private m_Warnings As Collection
Public Function RewriteFailure(ByVal fName As String, ByVal stage As Variant) As String
    InitGlobals
    Dim expr As CExpr: Set expr = EvalFunction(fName)
    Dim tpl As Object: Set tpl = LoadFormatTemplates()
    Dim body As String: body = RenderExprSymbolicLatex(expr, stage, tpl)
    RewriteFailure = ApplyQNamePrefixLatex(fName, body, tpl)
End Function

Public Function SubstituteFailure(ByVal fName As String, ByVal stage As Variant) As String
    InitGlobals
    Dim expr As CExpr: Set expr = EvalFunction(fName)
    Dim tpl As Object: Set tpl = LoadFormatTemplates()
    Dim body As String: body = RenderExprNumericLatex(expr, stage, tpl)
    SubstituteFailure = ApplyQNamePrefixLatex(fName, body, tpl)
End Function

Private Function ApplyQNamePrefixLatex(ByVal fName As String, ByVal body As String, ByVal tpl As Object) As String
    Dim prefixTpl As String
    prefixTpl = GetTplWarn(tpl, "Q_PREFIX_TEMPLATE", "Q_{ {FNAME} }\;=\;{BODY}")
    ApplyQNamePrefixLatex = ApplyTokens(prefixTpl, Array("FNAME", "BODY"), Array(EscapeLatexText(fName), body))
End Function

'===========================
' Symbolic rendering
'===========================

Private Function RenderExprSymbolicLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim tArr() As CTerm
    tArr = expr.GetTerms()
    If (Not Not tArr) = 0 Then
        RenderExprSymbolicLatex = GetTplWarn(tpl, "EMPTY_EXPR", "0")
        Exit Function
    End If

    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)

    Dim joinExpr As String: joinExpr = GetTplWarn(tpl, "SYM_EXPR_JOIN", " + ")
    Dim out As String: out = ""

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


'===========================
' Numeric rendering
'===========================

Private Function RenderExprNumericLatex(ByVal expr As CExpr, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim tArr() As CTerm
    tArr = expr.GetTerms()
    If (Not Not tArr) = 0 Then
        RenderExprNumericLatex = GetTplWarn(tpl, "EMPTY_EXPR", "0")
        Exit Function
    End If

    QuickSortCTermArray tArr, LBound(tArr), UBound(tArr)

    Dim joinExpr As String: joinExpr = GetTplWarn(tpl, "NUM_EXPR_JOIN", " + ")
    Dim out As String: out = ""

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



'===========================
' Split + helpers
'===========================

Private Sub SplitIDs_ByExtern(ByVal idsV As Variant, _
                             ByRef lamIDs() As Long, ByRef lamCount As Long, _
                             ByRef qIDs() As Long, ByRef qCount As Long)
    lamCount = 0: qCount = 0
    ReDim lamIDs(0 To 0)
    ReDim qIDs(0 To 0)

    Dim i As Long
    For i = LBound(idsV) To UBound(idsV)
        Dim id As Long: id = CLng(idsV(i))
        If Not m_ExternByID Is Nothing And m_ExternByID.Exists(id) Then
            If qCount = 0 Then
                qIDs(0) = id
            Else
                ReDim Preserve qIDs(0 To qCount)
                qIDs(qCount) = id
            End If
            qCount = qCount + 1
        Else
            If lamCount = 0 Then
                lamIDs(0) = id
            Else
                ReDim Preserve lamIDs(0 To lamCount)
                lamIDs(lamCount) = id
            End If
            lamCount = lamCount + 1
        End If
    Next i
End Sub

Private Function ComputeRTerm(ByVal lamCount As Long, ByRef qIDs() As Long, ByVal qCount As Long) As Long
    Dim sumRQ As Long: sumRQ = 0
    Dim i As Long
    For i = 0 To qCount - 1
        Dim id As Long: id = qIDs(i)
        Dim qi As Object: Set qi = m_ExternByID(id)
        sumRQ = sumRQ + CLng(qi("Order"))
    Next i
    ComputeRTerm = lamCount + sumRQ
End Function

Private Function ShouldSkipWi(ByVal lamCount As Long, ByRef qIDs() As Long, ByVal qCount As Long, ByVal stage As Variant) As Boolean
    If IsStageAll(stage) Then
        ShouldSkipWi = True
        Exit Function
    End If

    ' ??????????: ???? = ???? Q ??? ?, ? ??? ?????? ?? ??????
    If lamCount = 0 And qCount = 1 Then
        Dim id As Long: id = qIDs(0)
        Dim qi As Object: Set qi = m_ExternByID(id)
        If CBool(qi("HasStages")) Then
            ShouldSkipWi = True
            Exit Function
        End If
    End If

    ShouldSkipWi = False
End Function

Private Function GetWiValueSafe(ByVal r As Long, ByVal st As Long) As Double
    Dim maxR As Long: maxR = UBound(m_WiValues, 1)
    If r < 0 Or r > maxR Then Err.Raise vbObjectError + 880, "Writer", "Wi: r ??? ?????????: " & r
    GetWiValueSafe = m_WiValues(r, st)
End Function

Private Function EvalQTermNumeric(ByVal lamCount As Long, ByRef qIDs() As Long, ByVal qCount As Long, ByVal stage As Variant) As Double
    Dim isAll As Boolean: isAll = IsStageAll(stage)

    If qCount = 0 Then
        EvalQTermNumeric = 1#
        Exit Function
    End If

    If lamCount = 0 And qCount = 1 Then
        Dim id As Long: id = qIDs(0)
        Dim qi As Object: Set qi = m_ExternByID(id)

        If CBool(qi("HasStages")) Then
            If isAll Then
                EvalQTermNumeric = CDbl(qi("QAll"))
            Else
                Dim st As Long: st = CLng(stage)
                Dim arrV As Variant: arrV = qi("QStage")
                EvalQTermNumeric = CDbl(arrV(st))
            End If
        Else
            EvalQTermNumeric = CDbl(qi("QAll"))
        End If
        Exit Function
    End If

    Dim prod As Double: prod = 1#
    Dim i As Long
    For i = 0 To qCount - 1
        Dim qi2 As Object: Set qi2 = m_ExternByID(qIDs(i))
        prod = prod * CDbl(qi2("QAll"))
    Next i
    EvalQTermNumeric = prod
End Function

Private Function RenderLambdaProductByIDs(ByRef ids() As Long, ByVal cnt As Long, ByVal lamTpl As String, ByVal lamJoin As String) As String
    If cnt <= 0 Then RenderLambdaProductByIDs = "": Exit Function
    Dim s As String: s = ""
    Dim i As Long
    For i = 0 To cnt - 1
        Dim nm As String: nm = GetElementNameByID(ids(i))
        Dim one As String: one = ApplyTokens(lamTpl, Array("name", "id"), Array(EscapeLatexText(nm), CStr(ids(i))))
        If Len(s) > 0 Then s = s & lamJoin
        s = s & one
    Next i
    RenderLambdaProductByIDs = s
End Function

Private Function RenderQProductByIDs(ByRef ids() As Long, ByVal cnt As Long, ByVal stage As Variant, ByVal qTpl As String, ByVal qJoin As String) As String
    If cnt <= 0 Then RenderQProductByIDs = "": Exit Function

    Dim s As String: s = ""
    Dim i As Long
    Dim stText As String: stText = StageToText(stage)

    For i = 0 To cnt - 1
        Dim nm As String: nm = GetElementNameByID(ids(i))

        Dim basis As String
        basis = "t_{?}" ' ?? ????????? "?? ??? ?????"

        If Not IsStageAll(stage) Then
            If Not m_ExternByID Is Nothing And m_ExternByID.Exists(ids(i)) Then
                Dim qi As Object: Set qi = m_ExternByID(ids(i))
                If CBool(qi("HasStages")) Then
                    basis = "t_{" & stText & "}"
                End If
            End If
        End If

        Dim one As String
        one = ApplyTokens(qTpl, _
            Array("name", "id", "basis"), _
            Array(EscapeLatexText(nm), CStr(ids(i)), basis))

        If Len(s) > 0 Then s = s & qJoin
        s = s & one
    Next i

    RenderQProductByIDs = s
End Function

'===========================
' Templates + utils (??????? ?????????)
'===========================

Public Function LoadFormatTemplates() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("Q_PREFIX_TEMPLATE") = "Q_{ [[FNAME]] }\;=\;[[BODY]]"
    d("EMPTY_EXPR") = "0"
    
    d("SYM_EXPR_JOIN") = " + "
    d("SYM_TERM_TEMPLATE") = "[[MULT]][[WI]][[WI_MUL]][[LAMQPROD]][[TP]]"
    d("SYM_MULT_TEMPLATE") = "[[mult]]\,"
    d("SYM_WI_TEMPLATE") = "W_{ [[r]] }^{([[stage]])}"
    d("SYM_WI_MUL") = "\,\cdot\,"
    
    d("SYM_LAM_TEMPLATE") = "\lambda_{\text{[[name]]}}"
    d("SYM_LAM_JOIN") = "\cdot "
    
    d("SYM_Q_TEMPLATE") = "Q_{\text{[[name]]}}"
    d("SYM_Q_JOIN") = "\cdot "
    
    d("SYM_FACTOR_JOIN") = "\cdot "
    
    d("NUM_EXPR_JOIN") = " + "
    d("NUM_TERM_TEMPLATE") = "[[FACTORS]][[TP]]"
    d("NUM_FACTOR_JOIN") = "\,\cdot\,"
    
    d("NUM_PLAIN_MIN") = "0.001"
    d("NUM_PLAIN_MAX") = "1000"
    d("NUM_PLAIN_FMT") = "0.############"
    d("NUM_MANTISSA_FMT") = "0.#####"
    d("NUM_SCI_TEMPLATE") = "[[mant]]\cdot 10^{[[exp]]}"
    
    d("TP_SYM_1") = "\,t_{?}"
    d("TP_SYM_POW") = "\,t_{?}^{ [[r]] }"
    d("TP_NUM_1") = "\,\cdot\,[[tp]]"
    d("TP_NUM_POW") = "\,\cdot\,([[tp]])^{ [[r]] }"

    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Format")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim r As Long
        For r = 1 To lastRow
            Dim k As String: k = Trim$(CStr(ws.Cells(r, 1).value))
            Dim v As String: v = CStr(ws.Cells(r, 2).value)
            If Len(k) > 0 Then d(k) = v
        Next r
    End If

    Set LoadFormatTemplates = d
End Function

Public Function GetTplWarn(ByVal tpl As Object, ByVal key As String, ByVal defaultValue As String) As String
    If Not tpl Is Nothing Then
        If tpl.Exists(key) Then
            GetTplWarn = CStr(tpl(key))
            Exit Function
        End If
    End If

    ' ?????????????? ? ???????????
    Call DiagWarn(2001, "?? ?????? ?????? '" & key & "' ?? ????? Format. ????????? ??????: " & defaultValue)

    GetTplWarn = defaultValue
End Function


Public Function ApplyTokens(ByVal template As String, ByVal keys As Variant, ByVal values As Variant) As String
    Dim s As String: s = template
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        s = Replace(s, "[[" & CStr(keys(i)) & "]]", CStr(values(i)))
    Next i
    ApplyTokens = s
End Function


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
    s = Format$(v, "0.############")
    s = Replace(s, ",", ".")
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    If s = "-0" Then s = "0"
    TrimNumberSymbolic = s
End Function

Private Function RenderTpSymbolic(ByVal r As Long, ByVal tpl As Object) As String
    If r <= 0 Then
        RenderTpSymbolic = ""
    ElseIf r = 1 Then
        RenderTpSymbolic = GetTplWarn(tpl, "TP_SYM_1", "\,t_p")
    Else
        RenderTpSymbolic = ApplyTokens(GetTplWarn(tpl, "TP_SYM_POW", "\,t_p^{ {r} }"), Array("r"), Array(CStr(r)))
    End If
End Function

Private Function RenderTpNumeric(ByVal r As Long, ByVal tpl As Object) As String
    If r <= 0 Then RenderTpNumeric = "": Exit Function
    Dim tpStr As String: tpStr = FormatNumLatex(m_Tp, tpl)
    If r = 1 Then
        RenderTpNumeric = ApplyTokens(GetTplWarn(tpl, "TP_NUM_1", "\,\cdot\,{tp}"), Array("tp", "base"), Array(tpStr, tpStr))
    Else
        RenderTpNumeric = ApplyTokens(GetTplWarn(tpl, "TP_NUM_POW", "\,\cdot\,({tp})^{ {r} }"), Array("tp", "base", "r"), Array(tpStr, tpStr, CStr(r)))
    End If
End Function

Private Function FormatNumLatex(ByVal v As Double, ByVal tpl As Object) As String
    Dim plainMin As Double, plainMax As Double
    plainMin = CDblSafe(GetTplWarn(tpl, "NUM_PLAIN_MIN", "0.001"), 0.001)
    plainMax = CDblSafe(GetTplWarn(tpl, "NUM_PLAIN_MAX", "1000"), 1000#)

    If v = 0# Then
        FormatNumLatex = "0"
        Exit Function
    End If

    ' ? ???? ?????? ????????????? ?????, ?? ?? ?????? ?????? ???????????
    Dim av As Double
    av = Abs(v)

    ' ??????? (?? ???????) ??????
    If av >= plainMin And av < plainMax Then
        Dim s As String
        s = Format$(av, GetTplWarn(tpl, "NUM_PLAIN_FMT", "0.############"))
        ' VBA ????? ??????? ??????? ??? ?????????? ???????????
        s = Replace(s, ",", ".")
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
        s = Replace(s, ".", ",")
        FormatNumLatex = s
        Exit Function
    End If

    ' ??????? ??????: ???????? ??????????? ? ???????? 1 <= mant < 10
    Dim exp As Long
    exp = Int(Log(av) / Log(10#))   ' ?????: Int (floor), ? ?? Fix

    Dim mant As Double
    mant = av / (10# ^ exp)

    ' ??????? ?? ????????? [1, 10) ?? ?????? (?? ?????? ??????????? ????????)
    If mant >= 10# Then
        mant = mant / 10#
        exp = exp + 1
    ElseIf mant < 1# Then
        mant = mant * 10#
        exp = exp - 1
    End If

    Dim mantStr As String
    mantStr = Format$(mant, GetTplWarn(tpl, "NUM_MANTISSA_FMT", "0.#####"))
    mantStr = Replace(mantStr, ",", ".")
    If Right$(mantStr, 1) = "." Then mantStr = Left$(mantStr, Len(mantStr) - 1)
    mantStr = Replace(mantStr, ".", ",")

    ' ?????? ??????? ??????
    ' ??????: [[mant]]\cdot 10^{[[exp]]} (??? ???? ?????? ? {mant}/{exp})
    FormatNumLatex = ApplyTokens( _
        GetTplWarn(tpl, "NUM_SCI_TEMPLATE", "{mant}\cdot 10^{ {exp} }"), _
        Array("mant", "exp"), _
        Array(mantStr, CStr(exp)) _
    )
End Function

Private Function CDblSafe(ByVal s As String, ByVal defaultValue As Double) As Double
    On Error GoTo EH
    CDblSafe = CDbl(Replace(s, ".", ","))
    Exit Function
EH:
    CDblSafe = defaultValue
End Function

Private Function JoinNonEmpty(ByVal parts As Variant, ByVal delim As String) As String
    Dim i As Long, out As String
    out = ""
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = CStr(parts(i))
        If Len(p) > 0 Then
            If Len(out) > 0 Then out = out & delim
            out = out & p
        End If
    Next i
    JoinNonEmpty = out
End Function

' ?????????? ?????? ???? ??????? (??? ? ???? ? Writer), ?? ???????? ?????
Private Sub QuickSortCTermArray(ByRef arr() As CTerm, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    i = first: j = last
    Dim pivot As CTerm: Set pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While StrComp(arr(i).key, pivot.key, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(arr(j).key, pivot.key, vbTextCompare) > 0
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

Private Function StageToText(ByVal stage As Variant) As String
    If IsStageAll(stage) Then
        StageToText = "ALL"
        Exit Function
    End If

    Dim st As Long: st = CLng(stage)
    If st = 0 Then
        StageToText = "01"
    Else
        StageToText = CStr(st)
    End If
End Function



Private Sub WarningsReset()
    Set m_Warnings = New Collection
End Sub

Private Sub Warn(ByVal msg As String)
    If m_Warnings Is Nothing Then Set m_Warnings = New Collection
    m_Warnings.Add msg
End Sub

Private Function WarningsToText() As String
    Dim s As String, i As Long
    If m_Warnings Is Nothing Then Exit Function
    If m_Warnings.Count = 0 Then Exit Function

    s = "WARNINGS:" & vbCrLf
    For i = 1 To m_Warnings.Count
        s = s & " - " & CStr(m_Warnings(i)) & vbCrLf
    Next i
    WarningsToText = s
End Function


' ================================
' œ‡Ú˜ ‰Îˇ Writer.bas - ÔÓ‰‰ÂÊÍ‡ ÍÓÏÔ‡ÍÚÌ˚ı ÚÂÏÓ‚
' ================================
' «‡ÏÂÌËÚÂ ÙÛÌÍˆËË RenderOneCTermSymbolicLatex Ë RenderOneCTermNumericLatex

Private Function RenderOneCTermSymbolicLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    If Abs(t.Multiplier) < 0.0000000001 Then
        RenderOneCTermSymbolicLatex = ""
        Exit Function
    End If

    ' ===== Œ¡–¿¡Œ“ ¿  ŒÃœ¿ “Õ€’ “≈–ÃŒ¬ =====
    If t.TermType = 1 Then  ' ttCompact
        RenderOneCTermSymbolicLatex = RenderCompactTermSymbolic(t, stage, tpl)
        Exit Function
    End If
    
    ' ===== Œ¡–¿¡Œ“ ¿  ≈ÿ»–Œ¬¿ÕÕ€’ ‘”Õ ÷»… =====
    If t.TermType = 2 Then  ' ttCachedFunc
        RenderOneCTermSymbolicLatex = RenderCachedFuncSymbolic(t, stage, tpl)
        Exit Function
    End If

    ' ===== Œ¡€◊Õ¿ﬂ Œ¡–¿¡Œ“ ¿ =====
    Dim termTpl As String: termTpl = GetTplWarn(tpl, "SYM_TERM_TEMPLATE", "{MULT}{WI}{WI_MUL}{LAMQPROD}{TP}")
    Dim multTpl As String: multTpl = GetTplWarn(tpl, "SYM_MULT_TEMPLATE", "{mult}\,")
    Dim wiTpl As String: wiTpl = GetTplWarn(tpl, "SYM_WI_TEMPLATE", "W_{ {r} }^{({stage})}\,")
    Dim lamTpl As String: lamTpl = GetTplWarn(tpl, "SYM_LAM_TEMPLATE", "\lambda_{\text{{name}}}")
    Dim lamJoin As String: lamJoin = GetTplWarn(tpl, "SYM_LAM_JOIN", "\cdot ")
    Dim qTpl As String: qTpl = GetTplWarn(tpl, "SYM_Q_TEMPLATE", "Q_{ \text{{name}} }")
    Dim qJoin As String: qJoin = GetTplWarn(tpl, "SYM_Q_JOIN", "\cdot ")
    Dim factorJoin As String: factorJoin = GetTplWarn(tpl, "SYM_FACTOR_JOIN", "\cdot ")

    Dim multStr As String: multStr = ""
    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumberSymbolic(t.Multiplier)))
    End If

    Dim idsV As Variant
    idsV = t.FactorIDs
    If IsEmpty(idsV) Then
        RenderOneCTermSymbolicLatex = ""
        Exit Function
    End If

    Dim isAll As Boolean: isAll = IsStageAll(stage)
    Dim st As Long: If Not isAll Then st = CLng(stage)

    Dim lamIDs() As Long, lamCount As Long
    Dim qIDs() As Long, qCount As Long
    SplitIDs_ByExtern idsV, lamIDs, lamCount, qIDs, qCount

    Dim rTerm As Long
    rTerm = ComputeRTerm(lamCount, qIDs, qCount)

    Dim skipWi As Boolean
    skipWi = ShouldSkipWi(lamCount, qIDs, qCount, stage)
    
    Dim stText As String
    stText = StageToText(stage)

    Dim wiStr As String: wiStr = ""
    If (Not isAll) And (Not skipWi) Then
        Dim wiVal As Double
        wiVal = GetWiValueSafe(rTerm, st)
        If Abs(wiVal - 1#) > 0.0000000001 Then
            wiStr = ApplyTokens(wiTpl, Array("r", "stage"), Array(CStr(rTerm), stText))
        End If
    End If

    Dim wiMulStr As String
    If Len(wiStr) > 0 Then wiMulStr = GetTplWarn(tpl, "SYM_WI_MUL", "\,\cdot\,") Else wiMulStr = ""

    Dim lamProd As String
    lamProd = RenderLambdaProductByIDs(lamIDs, lamCount, lamTpl, lamJoin)

    Dim qProd As String
    qProd = RenderQProductByIDs(qIDs, qCount, stText, qTpl, qJoin)

    Dim lamQProd As String
    lamQProd = JoinNonEmpty(Array(lamProd, qProd), factorJoin)

    Dim tpStr As String
    tpStr = RenderTpSymbolic(lamCount, tpl)

    RenderOneCTermSymbolicLatex = ApplyTokens(termTpl, _
        Array("MULT", "WI", "WI_MUL", "LAMQPROD", "TP"), _
        Array(multStr, wiStr, wiMulStr, lamQProd, tpStr))
End Function

Private Function RenderOneCTermNumericLatex(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    If Abs(t.Multiplier) < 0.0000000001 Then
        RenderOneCTermNumericLatex = ""
        Exit Function
    End If

    ' ===== Œ¡–¿¡Œ“ ¿  ŒÃœ¿ “Õ€’ “≈–ÃŒ¬ =====
    If t.TermType = 1 Then  ' ttCompact
        RenderOneCTermNumericLatex = RenderCompactTermNumeric(t, stage, tpl)
        Exit Function
    End If
    
    ' ===== Œ¡–¿¡Œ“ ¿  ≈ÿ»–Œ¬¿ÕÕ€’ ‘”Õ ÷»… =====
    If t.TermType = 2 Then  ' ttCachedFunc
        RenderOneCTermNumericLatex = RenderCachedFuncNumeric(t, stage, tpl)
        Exit Function
    End If

    ' ===== Œ¡€◊Õ¿ﬂ Œ¡–¿¡Œ“ ¿ =====
    Dim factorJoin As String: factorJoin = GetTplWarn(tpl, "NUM_FACTOR_JOIN", "\,\cdot\,")
    Dim termTpl As String: termTpl = GetTplWarn(tpl, "NUM_TERM_TEMPLATE", "{FACTORS}{TP}")

    Dim factors As Collection: Set factors = New Collection

    Dim isAll As Boolean: isAll = IsStageAll(stage)
    Dim st As Long: If Not isAll Then st = CLng(stage)

    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        factors.Add FormatNumLatex(t.Multiplier, tpl)
    End If

    Dim idsV As Variant
    idsV = t.FactorIDs
    If IsEmpty(idsV) Then
        RenderOneCTermNumericLatex = ""
        Exit Function
    End If

    Dim lamIDs() As Long, lamCount As Long
    Dim qIDs() As Long, qCount As Long
    SplitIDs_ByExtern idsV, lamIDs, lamCount, qIDs, qCount

    Dim rTerm As Long
    rTerm = ComputeRTerm(lamCount, qIDs, qCount)

    Dim skipWi As Boolean
    skipWi = ShouldSkipWi(lamCount, qIDs, qCount, stage)

    If (Not isAll) And (Not skipWi) Then
        Dim wiVal As Double
        wiVal = GetWiValueSafe(rTerm, st)
        If Abs(wiVal - 1#) > 0.0000000001 Then
            factors.Add FormatNumLatex(wiVal, tpl)
        End If
    End If

    Dim i As Long
    For i = 0 To lamCount - 1
        factors.Add FormatNumLatex(m_LambdaValues(lamIDs(i)), tpl)
    Next i

    If qCount > 0 Then
        factors.Add FormatNumLatex(EvalQTermNumeric(lamCount, qIDs, qCount, stage), tpl)
    End If

    Dim factorsStr As String
    factorsStr = JoinCollection(factors, factorJoin)

    Dim tpStr As String
    tpStr = RenderTpNumeric(lamCount, tpl)

    RenderOneCTermNumericLatex = ApplyTokens(termTpl, Array("FACTORS", "TP"), Array(factorsStr, tpStr))
End Function

' ===== ÕŒ¬€≈ ‘”Õ ÷»» ƒÀﬂ  ŒÃœ¿ “Õ€’ “≈–ÃŒ¬ =====


Private Function RenderCompactTermSymbolic(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim r As Long: r = t.Order
    
    Dim multTpl As String: multTpl = GetTplWarn(tpl, "SYM_MULT_TEMPLATE", "{mult}\,")
    Dim wiTpl As String: wiTpl = GetTplWarn(tpl, "SYM_WI_TEMPLATE", "W_{ {r} }^{({stage})}\,")
    Dim factorJoin As String: factorJoin = "\cdot "
    
    Dim isAll As Boolean: isAll = IsStageAll(stage)
    Dim st As Long: If Not isAll Then st = CLng(stage)
    Dim stText As String: stText = StageToText(stage)
    
    ' Multiplier
    Dim multStr As String: multStr = ""
    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        multStr = ApplyTokens(multTpl, Array("mult"), Array(TrimNumberSymbolic(t.Multiplier)))
    End If
    
    ' Wi Œƒ»Õ –¿« ‰Îˇ ‚ÒÂ„Ó ÚÂÏ‡
    Dim wiStr As String: wiStr = ""
    If Not isAll Then
        Dim wiVal As Double
        wiVal = GetWiValueSafe(r, st)
        If Abs(wiVal - 1#) > 0.0000000001 Then
            wiStr = ApplyTokens(wiTpl, Array("r", "stage"), Array(CStr(r), stText))
        End If
    End If
    
    Dim wiMulStr As String
    If Len(wiStr) > 0 Then wiMulStr = "\,\cdot\," Else wiMulStr = ""
    
    ' ===== »—œ–¿¬À≈Õ»≈: –ÂÌ‰ÂËÏ Ù‡ÍÚÓ˚ Í‡Í ˜ËÒÚ˚Â ÒÛÏÏ˚ ? ¡≈« W Ë tp =====
    Dim factorsParts() As String
    ReDim factorsParts(0 To t.CompactFactors.Count - 1)
    
    Dim i As Long: i = 0
    Dim f As Variant
    For Each f In t.CompactFactors
        If TypeName(f) = "CExpr" Then
            Dim factorExpr As CExpr
            Set factorExpr = f
            ' –ÂÌ‰ÂËÏ ÚÓÎ¸ÍÓ ? ·ÂÁ W Ë tp
            factorsParts(i) = "(" & RenderPureLambdaSum(factorExpr, tpl) & ")"
        ElseIf TypeName(f) = "CTerm" Then
            Dim ft As CTerm
            Set ft = f
            If ft.TermType = 2 Then  ' ttCachedFunc
                factorsParts(i) = "Q_{\text{" & ft.FuncName & "}}"
            End If
        End If
        i = i + 1
    Next f
    
    Dim factorsStr As String
    factorsStr = Join(factorsParts, factorJoin)
    
    ' Tp
    Dim tpStr As String
    tpStr = RenderTpSymbolic(r, tpl)
    
    RenderCompactTermSymbolic = multStr & wiStr & wiMulStr & factorsStr & tpStr
End Function

' ===== ÕŒ¬¿ﬂ ‘”Õ ÷»ﬂ: –ÂÌ‰ÂËÌ„ ˜ËÒÚÓÈ ÒÛÏÏ˚ ? ¡≈« W Ë tp =====
Private Function RenderPureLambdaSum(ByVal expr As CExpr, ByVal tpl As Object) As String
    Dim terms() As CTerm
    terms = expr.GetTerms()
    
    On Error Resume Next
    Dim ub As Long, lb As Long
    lb = LBound(terms)
    ub = UBound(terms)
    On Error GoTo 0
    
    If Err.Number <> 0 Or ub < lb Then
        RenderPureLambdaSum = "0"
        Exit Function
    End If
    
    Dim lamTpl As String: lamTpl = GetTplWarn(tpl, "SYM_LAM_TEMPLATE", "\lambda_{\text{{name}}}")
    Dim joinStr As String: joinStr = " + "
    Dim multTpl As String: multTpl = GetTplWarn(tpl, "SYM_MULT_TEMPLATE", "{mult}\,")
    
    Dim parts() As String
    ReDim parts(0 To ub - lb)
    
    Dim i As Long, idx As Long
    idx = 0
    
    For i = lb To ub
        If Not terms(i) Is Nothing Then
            Dim ids() As Long
            ids = terms(i).FactorIDs
            
            If UBound(ids) = LBound(ids) Then
                Dim id As Long
                id = ids(LBound(ids))
                
                Dim nm As String
                If id <= UBound(m_IDToName) Then
                    nm = m_IDToName(id)
                Else
                    nm = "ID" & id
                End If
                
                Dim part As String
                part = ""
                
                ' Multiplier
                If Abs(terms(i).Multiplier - 1#) > 0.0000000001 Then
                    part = ApplyTokens(multTpl, Array("mult"), Array(TrimNumberSymbolic(terms(i).Multiplier)))
                End If
                
                ' Lambda
                part = part & ApplyTokens(lamTpl, Array("name"), Array(nm))
                
                parts(idx) = part
                idx = idx + 1
            End If
        End If
    Next i
    
    ' —ÍÎÂË‚‡ÂÏ
    Dim result As String
    result = ""
    For i = 0 To idx - 1
        If Len(parts(i)) > 0 Then
            If Len(result) > 0 Then result = result & joinStr
            result = result & parts(i)
        End If
    Next i
    
    RenderPureLambdaSum = result
End Function
Private Function RenderCompactTermNumeric(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    Dim r As Long: r = t.Order
    
    Dim factorJoin As String: factorJoin = GetTplWarn(tpl, "NUM_FACTOR_JOIN", "\,\cdot\,")
    
    Dim isAll As Boolean: isAll = IsStageAll(stage)
    Dim st As Long: If Not isAll Then st = CLng(stage)
    
    Dim factors As Collection
    Set factors = New Collection
    
    ' Multiplier
    If Abs(t.Multiplier - 1#) > 0.0000000001 Then
        factors.Add FormatNumLatex(t.Multiplier, tpl)
    End If
    
    ' Wi
    If Not isAll Then
        Dim wiVal As Double
        wiVal = GetWiValueSafe(r, st)
        If Abs(wiVal - 1#) > 0.0000000001 Then
            factors.Add FormatNumLatex(wiVal, tpl)
        End If
    End If
    
    ' ‘‡ÍÚÓ˚ - “ŒÀ‹ Œ ÓÌË ‚ ÒÍÓ·Í‡ı
    Dim f As Variant
    For Each f In t.CompactFactors
        If TypeName(f) = "CExpr" Then
            Dim factorExpr As CExpr
            Set factorExpr = f
            
            ' —ÍÓ·ÍË “ŒÀ‹ Œ ‚ÓÍÛ„ ÒÛÏÏ˚
            factors.Add "(" & RenderPureLambdaSumNumeric(factorExpr, tpl) & ")"
        ElseIf TypeName(f) = "CTerm" Then
            Dim ft As CTerm
            Set ft = f
            If ft.TermType = 2 Then  ' ttCachedFunc
                Dim funcVal As Double
                funcVal = ft.GetValueForOrder(r)
                ' Œ‰ÌÓ ˜ËÒÎÓ - ¡≈« ÒÍÓ·ÓÍ
                factors.Add FormatNumLatex(funcVal, tpl)
            End If
        End If
    Next f
    
    ' Tp - ‰Ó·‡‚ÎˇÂÏ ˜ÂÂÁ RenderTpNumeric (ÓÌ Ò‡Ï ‰Ó·‡‚ÎˇÂÚ cdot)
    Dim factorsStr As String
    factorsStr = JoinCollection(factors, factorJoin)
    
    Dim tpStr As String
    tpStr = RenderTpNumeric(r, tpl)
    
    RenderCompactTermNumeric = factorsStr & tpStr
End Function
Private Function RenderPureLambdaSumNumeric(ByVal expr As CExpr, ByVal tpl As Object) As String
    Dim terms() As CTerm
    terms = expr.GetTerms()
    
    On Error Resume Next
    Dim ub As Long, lb As Long
    lb = LBound(terms)
    ub = UBound(terms)
    On Error GoTo 0
    
    If Err.Number <> 0 Or ub < lb Then
        RenderPureLambdaSumNumeric = "0"
        Exit Function
    End If
    
    Dim joinStr As String: joinStr = " + "
    
    Dim parts() As String
    ReDim parts(0 To ub - lb)
    
    Dim i As Long, idx As Long
    idx = 0
    
    For i = lb To ub
        If Not terms(i) Is Nothing Then
            Dim ids() As Long
            ids = terms(i).FactorIDs
            
            If UBound(ids) = LBound(ids) Then
                Dim id As Long
                id = ids(LBound(ids))
                
                If id <= UBound(m_LambdaValues) Then
                    Dim val As Double
                    val = m_LambdaValues(id) * terms(i).Multiplier
                    parts(idx) = FormatNumLatex(val, tpl)
                    idx = idx + 1
                End If
            End If
        End If
    Next i
    
    ' —ÍÎÂË‚‡ÂÏ
    Dim result As String
    result = ""
    For i = 0 To idx - 1
        If Len(parts(i)) > 0 Then
            If Len(result) > 0 Then result = result & joinStr
            result = result & parts(i)
        End If
    Next i
    
    RenderPureLambdaSumNumeric = result
End Function

' ===== ????? ???????: ?????????? ????? ? =====
Private Function CalcPureLambdaSum(ByVal expr As CExpr) As Double
    Dim terms() As CTerm
    terms = expr.GetTerms()
    
    On Error Resume Next
    Dim ub As Long, lb As Long
    lb = LBound(terms)
    ub = UBound(terms)
    On Error GoTo 0
    
    If Err.Number <> 0 Or ub < lb Then
        CalcPureLambdaSum = 0#
        Exit Function
    End If
    
    Dim sumVal As Double
    sumVal = 0#
    
    Dim i As Long
    For i = lb To ub
        If Not terms(i) Is Nothing Then
            Dim ids() As Long
            ids = terms(i).FactorIDs
            
            If UBound(ids) = LBound(ids) Then
                Dim id As Long
                id = ids(LBound(ids))
                
                If id <= UBound(m_LambdaValues) Then
                    sumVal = sumVal + m_LambdaValues(id) * terms(i).Multiplier
                End If
            End If
        End If
    Next i
    
    CalcPureLambdaSum = sumVal
End Function

Private Function RenderCachedFuncSymbolic(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    ' —ËÏ‚ÓÎ¸Ì˚È ‚˚‚Ó‰ ÍÂ¯ËÓ‚‡ÌÌÓÈ ÙÛÌÍˆËË
    RenderCachedFuncSymbolic = "Q_{\text{" & t.FuncName & "}}"
End Function

Private Function RenderCachedFuncNumeric(ByVal t As CTerm, ByVal stage As Variant, ByVal tpl As Object) As String
    ' ◊ËÒÎÂÌÌ˚È ‚˚‚Ó‰ ÍÂ¯ËÓ‚‡ÌÌÓÈ ÙÛÌÍˆËË
    Dim isAll As Boolean: isAll = IsStageAll(stage)
    Dim st As Long: If Not isAll Then st = CLng(stage)
    
    Dim orderVec As Object
    Set orderVec = GetOrComputeOrderVector(t.FuncName)
    
    Dim total As Double: total = 0#
    Dim k As Variant
    For Each k In orderVec.keys
        Dim r As Long: r = CLng(k)
        Dim wiVal As Double
        If Not isAll Then
            wiVal = GetWiValueSafe(r, st)
        Else
            wiVal = 1#
        End If
        Dim tpPow As Double: tpPow = m_Tp ^ r
        total = total + t.Multiplier * wiVal * tpPow * CDbl(orderVec(k))
    Next k
    
    RenderCachedFuncNumeric = FormatNumLatex(total, tpl)
End Function
