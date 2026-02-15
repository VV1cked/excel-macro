Attribute VB_Name = "Enums"
' Типы термов
Public Enum TermType
    ttNormal = 0        ' обычный терм (список IDs)
    ttCompact = 1       ' компактный терм (список факторов)
    ttCachedFunc = 2    ' кешированная функция (вектор по Order)
End Enum

