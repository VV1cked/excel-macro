Attribute VB_Name = "CalcUtils"
Option Explicit

' ====== ”ниверсальный парсер числа ======
' ѕоддерживает:
'   - дес€тичную зап€тую (0,123)
'   - scientific формат 1e-5 / 1E-5
'   - пробелы внутри (будут удалены)
Public Function ParseDouble(ByVal v As Variant, ByVal context As String) As Double
    Dim s As String
    s = Trim$(CStr(v))

    ' убираем пробелы
    s = Replace(s, " ", "")

    ' дес€тичный разделитель -> точка
    s = Replace(s, ",", ".")

    ' scientific: разрешаем 'e' и 'E'
    ' (в некоторых VBA/локал€х Val надЄжнее работает с 'E')
    s = Replace(s, "e", "E")

    If s <> "" Then
        ' Val понимает числа вида 1E-5, 2.3E+4 и т.п.
        ParseDouble = val(s)
    Else
        Err.Raise 996, , "Ќевозможно преобразовать в число: " & context & " = '" & CStr(v) & "'"
    End If
End Function

Public Function TryGetBounds(ByRef arr As Variant, ByRef lb As Long, ByRef ub As Long) As Boolean
    ' –аботает и дл€ массивов объектов (CTerm), и дл€ массивов Long, и дл€ Variant-массивов
    On Error GoTo Fail
    If IsArray(arr) = False Then GoTo Fail

    Err.Clear
    lb = LBound(arr)
    ub = UBound(arr)

    If ub < lb Then GoTo Fail
    TryGetBounds = True
    Exit Function

Fail:
    TryGetBounds = False
End Function
