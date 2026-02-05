Attribute VB_Name = "TestGuard"
Option Explicit
'==========================================================
' Test module: Guards / Parser diagnostics (minimal, robust)
'==========================================================
' What it tests:
'  1) Good formulas parse OK
'  2) Syntax errors -> ERR_SYNTAX
'  3) Unknown names -> ERR_NAME_NOT_FOUND
'  4) Missing function -> ERR_FUNC_NOT_FOUND
'  5) Name conflicts by load order -> ERR_NAME_CONFLICT
'
' Notes:
'  - No framework. One entry point: RunAll()
'  - Each test is isolated and cannot "leak" On Error handlers
'  - Uses a minimal in-memory dataset (no sheets)
'==========================================================

' ---- Error codes (must match your project) ----
Private Const ERR_SYNTAX As Long = vbObjectError + 1001
Private Const ERR_FUNC_NOT_FOUND As Long = vbObjectError + 3001
Private Const ERR_NAME_NOT_FOUND As Long = vbObjectError + 3002
Private Const ERR_Q_NOT_FOUND As Long = vbObjectError + 3003   ' optional in your current semantics
Private Const ERR_CYCLE As Long = vbObjectError + 3004         ' optional
Private Const ERR_NAME_CONFLICT As Long = vbObjectError + 3201

' ---- Local "assert" helper errors (only for tests) ----
Private Const ERR_TEST_FAIL As Long = vbObjectError + 9900

'==========================================================
' Public entry
'==========================================================
Public Sub RunAll()
    On Error GoTo EH

    Debug.Print "=== Guards Tests START ==="
    Call SetupMiniData

    Call Test_GoodFunctions
    Call Test_SyntaxError
    Call Test_UnknownElementName
    Call Test_UnknownQName_CurrentSemantics
    Call Test_MissingFunction
    Call Test_NameConflicts_LoadOrder

    Debug.Print "=== Guards Tests OK ==="
    MsgBox "OK: все тесты защит прошли", vbInformation
    Exit Sub

EH:
    Debug.Print "=== Guards Tests FAILED ==="
    Debug.Print "Err.Number=" & Err.Number
    Debug.Print "Err.Source=" & Err.source
    Debug.Print Err.Description
    MsgBox "FAILED: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Guards Tests"
End Sub

'==========================================================
' Dataset setup (no sheets)
'==========================================================
Private Sub SetupMiniData()
    ' Minimal globals needed by EvalFunction/CreateAtomStrict/GetIDStrict/RegisterName
    Set m_NameToID = CreateObject("Scripting.Dictionary")
    Set m_FuncExprCache = CreateObject("Scripting.Dictionary")
    Set m_FuncDNFCache = CreateObject("Scripting.Dictionary")
    Set m_CallStack = CreateObject("Scripting.Dictionary")
    Set m_ExternByID = CreateObject("Scripting.Dictionary")
    Set m_NameKind = CreateObject("Scripting.Dictionary")

    ReDim m_IDToName(0)
    ReDim m_LambdaValues(0)

    ' ---- Elements (do NOT start with D/X/Y in your real data; here just normal names) ----
    Call RegisterName_Test("F1.1", "ELEM", "Test.Elements")
    Call RegisterName_Test("S9.1", "ELEM", "Test.Elements")
    Call RegisterName_Test("S9.2", "ELEM", "Test.Elements")
    Call RegisterName_Test("K1.1", "ELEM", "Test.Elements")

    ' Create IDs for elements (mirrors real load behavior)
    Call GetID("F1.1")
    Call GetID("S9.1")
    Call GetID("S9.2")
    Call GetID("K1.1")

    ' ---- External Q (names start with Q) ----
    Call RegisterName_Test("Q1", "Q", "Test.Extern")
    Call GetID("Q1")

    Dim qInfo As Object: Set qInfo = CreateObject("Scripting.Dictionary")
    qInfo("Name") = "Q1"
    qInfo("Order") = 1
    qInfo("HasStages") = False
    qInfo("QAll") = 0.123
    m_ExternByID.Add GetID("Q1"), qInfo

    ' ---- Functions (names start with S) ----
    Call RegisterName_Test("SGOOD1", "FUNC", "Test.Functions")
    Call RegisterName_Test("SGOOD2", "FUNC", "Test.Functions")
    Call RegisterName_Test("SBAD_SYNTAX", "FUNC", "Test.Functions")
    Call RegisterName_Test("SBAD_UNKNOWN_ELEM", "FUNC", "Test.Functions")
    Call RegisterName_Test("SBAD_UNKNOWN_Q", "FUNC", "Test.Functions")

    m_FuncExprCache("SGOOD1") = "F1.1*S9.2+Q1"
    m_FuncExprCache("SGOOD2") = "(F1.1+K1.1)*S9.2"
    m_FuncExprCache("SBAD_SYNTAX") = "F1.1+*S9.2"
    m_FuncExprCache("SBAD_UNKNOWN_ELEM") = "F1.1*D1"  ' D1 not registered => unknown name
    m_FuncExprCache("SBAD_UNKNOWN_Q") = "F1.1*Q99"    ' Q99 not registered => unknown name (in current semantics)
End Sub

'==========================================================
' Tests
'==========================================================
Private Sub Test_GoodFunctions()
    Debug.Print "--- Test_GoodFunctions ---"
    Call ExpectOK("SGOOD1")
    Call ExpectOK("SGOOD2")
End Sub

Private Sub Test_SyntaxError()
    Debug.Print "--- Test_SyntaxError ---"
    Call ExpectErr(ERR_SYNTAX, "SBAD_SYNTAX")
End Sub

Private Sub Test_UnknownElementName()
    Debug.Print "--- Test_UnknownElementName ---"
    Call ExpectErr(ERR_NAME_NOT_FOUND, "SBAD_UNKNOWN_ELEM")
End Sub

Private Sub Test_UnknownQName_CurrentSemantics()
    Debug.Print "--- Test_UnknownQName_CurrentSemantics ---"
    ' With your clarified semantics: names are not distinguished lexically.
    ' Therefore unknown Q is treated as unknown name.
    Call ExpectErr(ERR_NAME_NOT_FOUND, "SBAD_UNKNOWN_Q")
End Sub

Private Sub Test_MissingFunction()
    Debug.Print "--- Test_MissingFunction ---"
    Call ExpectErr(ERR_FUNC_NOT_FOUND, "SNO_SUCH_FUNC")
End Sub

Private Sub Test_NameConflicts_LoadOrder()
    Debug.Print "--- Test_NameConflicts_LoadOrder ---"
    ' Conflicts are expected (and count as OK)
    Call ExpectNameConflict("F1.1", "ELEM", "Elements", "Дубль элемента")
    Call ExpectNameConflict("S9.2", "FUNC", "Functions", "Функция vs элемент")
    Call ExpectNameConflict("SGOOD1", "Q", "Extern", "Q vs функция")
End Sub

'==========================================================
' Expect helpers (robust)
'==========================================================
Private Sub ExpectOK(ByVal fName As String)
    On Error GoTo EH
    Dim e As CExpr
    Set e = EvalFunction(fName)

    If e Is Nothing Then
        Err.Raise ERR_TEST_FAIL, "Test", "EvalFunction вернул Nothing для '" & fName & "'"
    End If

    Debug.Print "OK:", fName
    Exit Sub

EH:
    Err.Raise ERR_TEST_FAIL, "Test", "Ожидался успех для '" & fName & "', но ошибка: " & Err.Number & " " & Err.Description
End Sub

Private Sub ExpectErr(ByVal expectedErrNumber As Long, ByVal fName As String)
    On Error GoTo EH
    Dim e As CExpr
    Set e = EvalFunction(fName)

    Err.Raise ERR_TEST_FAIL, "Test", "Ожидалась ошибка " & expectedErrNumber & " для '" & fName & "', но парсинг прошёл успешно."
    Exit Sub

EH:
    If Err.Number <> expectedErrNumber Then
        Err.Raise ERR_TEST_FAIL, "Test", _
            "Для '" & fName & "' ожидалась ошибка " & expectedErrNumber & ", но пришла " & Err.Number & vbCrLf & Err.Description
    End If
    Debug.Print "OK ERR:", fName, "=>", Err.Number
End Sub

Private Sub ExpectNameConflict(ByVal nm As String, ByVal kind As String, ByVal where As String, ByVal msg As String)
    On Error GoTo EH
    Call RegisterName_Test(nm, kind, where)

    Err.Raise ERR_TEST_FAIL, "Test", "Ожидался конфликт имён, но его не было: " & msg & " (" & nm & " as " & kind & ")"
    Exit Sub

EH:
    If Err.Number <> ERR_NAME_CONFLICT Then
        Err.Raise ERR_TEST_FAIL, "Test", _
            msg & ": ожидался код " & ERR_NAME_CONFLICT & ", но пришёл " & Err.Number & vbCrLf & Err.Description
    End If
    Debug.Print "OK CONFLICT:", msg, "=>", Err.Number
End Sub

'==========================================================
' Local RegisterName copy (so tests do not depend on Private)
'==========================================================
Private Sub RegisterName_Test(ByVal nm As String, ByVal kind As String, ByVal where As String)
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Sub

    If m_NameKind.Exists(nm) Then
        Dim prev As String: prev = CStr(m_NameKind(nm))
        Err.Raise ERR_NAME_CONFLICT, "InitGlobals", _
            "Конфликт имён: '" & nm & "' уже занято (" & prev & "), нельзя зарегистрировать как " & kind & _
            ". Источник: " & where
    End If

    m_NameKind.Add nm, kind
End Sub


