Attribute VB_Name = "CalcUtils"
Option Explicit

' ====== Универсальный парсер числа ======
' Поддерживает:
'   - десятичную запятую (0,123)
'   - scientific формат 1e-5 / 1E-5
'   - пробелы внутри (будут удалены)
Public Function ParseDouble(ByVal v As Variant, ByVal context As String) As Double
    Dim s As String
    s = Trim$(CStr(v))

    ' убираем пробелы
    s = Replace(s, " ", "")

    ' десятичный разделитель -> точка
    s = Replace(s, ",", ".")

    ' scientific: разрешаем 'e' и 'E'
    ' (в некоторых VBA/локалях Val надёжнее работает с 'E')
    s = Replace(s, "e", "E")

    If s <> "" Then
        ' Val понимает числа вида 1E-5, 2.3E+4 и т.п.
        ParseDouble = Val(s)
    Else
        Err.Raise 996, , "Невозможно преобразовать в число: " & context & " = '" & CStr(v) & "'"
    End If
End Function

