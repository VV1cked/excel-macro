Attribute VB_Name = "CalcUtils"
Option Explicit


' ====== Преобразование строки в число с поддержкой "e" и локализации ======
Public Function ParseDouble(ByVal v As Variant, ByVal context As String) As Double
    Dim s As String
    s = Trim(CStr(v))
    
    ' Убираем пробелы
    s = Replace(s, " ", "")
    ' Заменяем запятую на точку для десятичных
    s = Replace(s, ",", ".")
    
    ' Попытка через Val (работает с экспонентой)
    If s <> "" Then
        ParseDouble = Val(s)
    Else
        Err.Raise 996, , "Невозможно преобразовать в число: " & context & " = '" & v & "'"
    End If
End Function

