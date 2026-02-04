VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TermsDict As Object

Private Sub Class_Initialize()
    Set m_TermsDict = CreateObject("Scripting.Dictionary")
End Sub

' Добавить терм
Public Sub AddTerm(ByRef t As CTerm)
    If m_TermsDict.Exists(t.key) Then
        Dim existing As CTerm: Set existing = m_TermsDict(t.key)
        existing.Multiplier = existing.Multiplier + t.Multiplier
        If existing.Multiplier = 0 Then m_TermsDict.Remove t.key
    Else
        If t.Multiplier <> 0 Then Set m_TermsDict(t.key) = t
    End If
End Sub

' Получить массив терминов
Public Function GetTerms() As CTerm()
    Dim res() As CTerm
    If m_TermsDict.Count = 0 Then
        GetTerms = res: Exit Function
    End If
    ReDim res(0 To m_TermsDict.Count - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In m_TermsDict.keys
        Set res(i) = m_TermsDict(k)
        i = i + 1
    Next
    GetTerms = res
End Function

