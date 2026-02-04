VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FactorIDs() As Long
Public Multiplier As Double
Public Order As Long
Public key As String

' Инициализация терма
Public Sub Init(ByRef ids() As Long, ByVal mult As Double, ByVal sKey As String)
    Dim i As Long
    If (Not Not ids) = 0 Then Exit Sub
    ReDim m_FactorIDs(LBound(ids) To UBound(ids))
    For i = LBound(ids) To UBound(ids)
        m_FactorIDs(i) = ids(i)
    Next i
    Multiplier = mult
    Order = UBound(ids) - LBound(ids) + 1
    key = sKey
End Sub

' Получить ID факторов
Public Property Get FactorIDs() As Long()
    FactorIDs = m_FactorIDs
End Property

' Установить ID факторов
Public Property Let FactorIDs(ByRef ids() As Long)
    m_FactorIDs = ids
    On Error Resume Next
    Order = UBound(ids) - LBound(ids) + 1
    On Error GoTo 0
End Property

