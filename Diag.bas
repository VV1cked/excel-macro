Attribute VB_Name = "Diag"
Option Explicit

Public m_Warnings As Collection

Public Sub DiagInit()
    Set m_Warnings = New Collection
End Sub

Public Sub DiagWarn(ByVal code As Long, ByVal msg As String)
    If m_Warnings Is Nothing Then DiagInit
    m_Warnings.Add "W" & CStr(code) & ": " & msg
End Sub

Public Sub ShowWarningsIfAny(Optional ByVal title As String = "Предупреждения")
    If m_Warnings Is Nothing Then Exit Sub
    If m_Warnings.Count = 0 Then Exit Sub

    Dim i As Long, s As String
    For i = 1 To m_Warnings.Count
        s = s & m_Warnings(i) & vbCrLf
    Next i
    MsgBox s, vbExclamation, title
End Sub

