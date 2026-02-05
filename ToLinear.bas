Attribute VB_Name = "ToLinear"
Option Explicit
'=========================================================
' Module: LatexWordLinearTranslator
' Translate small LaTeX subset -> Word Linear (UnicodeMath)
' Fixes: Word swallowing "(t_4)" into subscript for Q_"Q2"(t_4)
' by inserting U+2061 FUNCTION APPLICATION between subscript and "("
'=========================================================

' --- Public API ------------------------------------------------------------

Public Function LatexToWordLinear_Full(ByVal latex As String) As String
    Dim s As String
    s = latex

    ' 1) Remove LaTeX spacing commands
    s = NormalizeLatexSpaces(s)

    ' 2) Symbols / operators (use ChrW to avoid "?")
    s = Replace(s, "\lambda", ChrW(&H3BB))      ' ?
    s = Replace(s, "\cdot", ChrW(&H22C5))       ' ?

    ' 3) \text{...} -> "..."
    s = ReplaceTextMacros(s)

    ' 4) Convert scripts _{...}, ^{...}, _x, ^x -> _(...), ^(...)
    s = ConvertScripts(s)

    ' 5) Cleanup
    s = CleanupLinear(s)

    ' 6) Normalize subscripts: _("Q2")->_"Q2", _(SYS6)->_SYS6, etc.
    s = NormalizeWordLinearSubscripts(s)

    ' 7) CRITICAL: prevent Q_"Q2"(...) from swallowing "(...)" into subscript
    ' Insert U+2061 FUNCTION APPLICATION between subscripted base and "("
    s = InsertFuncAppAfterSubscript(s)

    LatexToWordLinear_Full = s
End Function


' --- Phase 1: cleanup ------------------------------------------------------

Private Function NormalizeLatexSpaces(ByVal s As String) As String
    s = Replace(s, "\;", " ")
    s = Replace(s, "\,", " ")
    s = Replace(s, "\ ", " ")
    NormalizeLatexSpaces = CollapseSpaces(s)
End Function

Private Function CleanupLinear(ByVal s As String) As String
    s = CollapseSpaces(s)

    s = Replace(s, "{ ", "{")
    s = Replace(s, " }", "}")
    s = Replace(s, "( ", "(")
    s = Replace(s, " )", ")")

    ' Optional tightening
    s = Replace(s, " " & ChrW(&H22C5) & " ", ChrW(&H22C5))
    s = Replace(s, " + ", "+")
    s = Replace(s, " = ", "=")

    CleanupLinear = s
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = s
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = Trim$(t)
End Function


' --- Phase 2: \text{...} handling -----------------------------------------

Private Function ReplaceTextMacros(ByVal s As String) As String
    Dim i As Long
    i = InStr(1, s, "\text{", vbBinaryCompare)

    Do While i > 0
        Dim braceOpen As Long
        braceOpen = i + Len("\text") ' should point to "{"
        If braceOpen <= Len(s) And Mid$(s, braceOpen, 1) = "{" Then
            Dim braceClose As Long
            braceClose = FindMatchingBrace(s, braceOpen)
            If braceClose = 0 Then Exit Do

            Dim content As String
            content = Mid$(s, braceOpen + 1, braceClose - braceOpen - 1)
            content = Trim$(content)

            s = Left$(s, i - 1) & """" & content & """" & Mid$(s, braceClose + 1)
            i = InStr(i + 1, s, "\text{", vbBinaryCompare)
        Else
            Exit Do
        End If
    Loop

    ReplaceTextMacros = s
End Function

Private Function FindMatchingBrace(ByVal s As String, ByVal openPos As Long) As Long
    If openPos < 1 Or openPos > Len(s) Then Exit Function
    If Mid$(s, openPos, 1) <> "{" Then Exit Function

    Dim depth As Long: depth = 0
    Dim i As Long
    For i = openPos To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindClosingQuote(ByVal s As String, ByVal openQuotePos As Long) As Long
    If openQuotePos < 1 Or openQuotePos > Len(s) Then Exit Function
    If Mid$(s, openQuotePos, 1) <> """" Then Exit Function

    Dim i As Long
    For i = openQuotePos + 1 To Len(s)
        If Mid$(s, i, 1) = """" Then
            FindClosingQuote = i
            Exit Function
        End If
    Next i
End Function


' --- Phase 3: script conversion -------------------------------------------

Private Function ConvertScripts(ByVal s As String) As String
    Dim i As Long: i = 1
    Do While i <= Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)

        If ch = "_" Or ch = "^" Then
            Dim j As Long: j = i + 1
            Do While j <= Len(s) And Mid$(s, j, 1) = " "
                j = j + 1
            Loop
            If j > Len(s) Then Exit Do

            Dim nextCh As String: nextCh = Mid$(s, j, 1)

            If nextCh = "{" Then
                Dim endBrace As Long: endBrace = FindMatchingBrace(s, j)
                If endBrace = 0 Then Exit Do

                Dim content As String
                content = Trim$(Mid$(s, j + 1, endBrace - j - 1))

                Dim repl As String
                repl = ch & "(" & content & ")"
                s = Left$(s, i - 1) & repl & Mid$(s, endBrace + 1)
                i = i + Len(repl)
                GoTo ContinueLoop

            ElseIf nextCh = """" Then
                Dim endQ As Long: endQ = FindClosingQuote(s, j)
                If endQ = 0 Then Exit Do

                Dim qContent As String
                qContent = Mid$(s, j, endQ - j + 1)

                Dim replQ As String
                replQ = ch & "(" & qContent & ")"
                s = Left$(s, i - 1) & replQ & Mid$(s, endQ + 1)
                i = i + Len(replQ)
                GoTo ContinueLoop

            Else
                ' single char token
                Dim repl1 As String
                repl1 = ch & "(" & nextCh & ")"
                s = Left$(s, i - 1) & repl1 & Mid$(s, j + 1)
                i = i + Len(repl1)
                GoTo ContinueLoop
            End If
        End If

        i = i + 1
ContinueLoop:
    Loop

    ConvertScripts = s
End Function


' --- Phase 4: normalization / function application fix ---------------------

Private Function NormalizeWordLinearSubscripts(ByVal s As String) As String
    ' _("Q2") -> _"Q2"
    s = ReReplace(s, "_\(\s*""([^""]+)""\s*\)", "_""$1""")

    ' _(SYS6) -> _SYS6, _(4)->_4, _(я)->_я, etc.
    s = ReReplace(s, "_\(\s*([A-Za-z0-9._Р-пр-џЈИ]+)\s*\)", "_$1")

    NormalizeWordLinearSubscripts = s
End Function

Private Function InsertFuncAppAfterSubscript(ByVal s As String) As String
    ' Insert FUNCTION APPLICATION (U+2061) between a subscripted base and "(".
    ' This prevents: Q_"Q2"(t_4) from being parsed as subscript content.
    Dim fa As String
    fa = ChrW(&H2061) ' ? (usually invisible in Professional)

    ' _"Q2"(  -> _"Q2"?(
    s = ReReplace(s, "_""([^""]+)""\(", "_""$1""" & fa & "(")

    ' _SYS6(  -> _SYS6?(
    s = ReReplace(s, "_([A-Za-z0-9._Р-пр-џЈИ]+)\(", "_$1" & fa & "(")

    InsertFuncAppAfterSubscript = s
End Function

Private Function ReReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Multiline = True
    re.IgnoreCase = False
    re.pattern = pattern
    ReReplace = re.Replace(text, replacement)
End Function



Public Sub CopySelectedCell_AsWordLinear()
    Dim rng As Range
    Set rng = Selection

    If rng Is Nothing Or rng.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 5000, , "Expected single selected cell."
    End If

    Dim out As String
    out = LatexToWordLinear_Full(CStr(rng.Value2))

    ' Fix Word-linear quirks like ^(()2), ^(()(2)), ^((2))
    out = SanitizeWordLinear(out)

    Clipboard out
End Sub

Public Function Clipboard$(Optional ByVal s As String = vbNullString)
    Dim html As Object
    Set html = CreateObject("htmlfile")

    If Len(s) > 0 Then
        html.parentWindow.clipboardData.setData "text", s
    Else
        Clipboard = html.parentWindow.clipboardData.GetData("text")
    End If
End Function

'=========================================================
' Post-normalization for Word Linear (UnicodeMath)
' - remove empty group "()" right after "^("
' - collapse ^((TOKEN)) -> ^(TOKEN) for simple tokens
'=========================================================
Private Function SanitizeWordLinear(ByVal s As String) As String
    ' 1) Remove empty group after ^(
    '    ^(()2)      -> ^(2)
    '    ^(()(2))    -> ^((2))   (then step 2 collapses further)
    s = ReReplace(s, "\^\(\(\)\s*", "^(")

    ' 2) Collapse double parentheses around a simple token:
    '    ^((2))      -> ^(2)
    '    ^((4))      -> ^(4)
    '    ^((я))      -> ^(я)
    '    ^((SYS6))   -> ^(SYS6)
    ' Token set: letters/digits/dot/underscore + Cyrillic incl. ЈИ
    s = ReReplace(s, "\^\(\(\s*([A-Za-z0-9._Р-пр-џЈИ]+)\s*\)\)", "^($1)")

    SanitizeWordLinear = s
End Function


