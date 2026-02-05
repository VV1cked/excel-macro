Attribute VB_Name = "Constants"
Option Explicit

'Option Explicit

' ====== Листы ======
Public Const SHEET_ELEMENTS As String = "Elements"
Public Const SHEET_FUNCTIONS As String = "Functions"
Public Const SHEET_WI As String = "Wi"

' [NEW] Лист заранее заданных вероятностей подсистем (Имя | Q | Order)
Public Const SHEET_PRECALC As String = "Precalc"

' ====== Диапазоны ======
Public Const RANGE_ELEMENTS_START As String = "A2"
Public Const RANGE_ELEMENTS_COL_NAME As Long = 1
Public Const RANGE_ELEMENTS_COL_LAMBDA As Long = 2

Public Const RANGE_FUNCTIONS_START As String = "A2"
Public Const RANGE_FUNCTIONS_COL_NAME As Long = 1
Public Const RANGE_FUNCTIONS_COL_EXPR As Long = 2

Public Const RANGE_WI_START As String = "A2"
Public Const RANGE_WI_COL_ORDER As Long = 1
Public Const RANGE_WI_COL_VALUE As Long = 2

' [NEW] Precalc: A=Name, B=Q (1 или 13 чисел в ячейке), C=Order (если пусто => 1)
Public Const RANGE_PRECALC_START As String = "A2"
Public Const RANGE_PRECALC_COL_NAME As Long = 1
Public Const RANGE_PRECALC_COL_Q As Long = 2
Public Const RANGE_PRECALC_COL_ORDER As Long = 3

' ====== Символы ======
Public Const CH_LPAREN As String = "("
Public Const CH_RPAREN As String = ")"
Public Const CH_AND As String = "*"
Public Const CH_OR As String = "+"

' ====== Этапы ======
Public Const STAGE_COUNT As Long = 13     ' Stage0..Stage12
Public Const STAGE_MIN As Long = 0
Public Const STAGE_MAX As Long = 12
Public Const STAGE_ALL As String = "ALL"

' [NEW] Разделители списка чисел в одной ячейке Q.
' ВАЖНО: запятая — десятичный разделитель, поэтому НЕ используем её как разделитель списка.
Public Const Q_LIST_SEPARATORS As String = ";" & vbTab & vbCr & vbLf

' ====== Пределы ======
' Текущий код режет умножение термов по R_MAX (MultiplyExpr) и Wi чтение (LoadWi).
' Пока оставляем как есть (4), чтобы не ломать текущую модель.
' На шаге интеграции Q мы пересмотрим этот механизм.
Public Const R_MAX As Long = 4

Public Const WI_COL_R As String = "A"
Public Const WI_COL_MAX As String = "O"
