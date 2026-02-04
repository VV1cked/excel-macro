Attribute VB_Name = "Constants"
Option Explicit

' ====== Листы ======
Public Const SHEET_ELEMENTS As String = "Elements"
Public Const SHEET_FUNCTIONS As String = "Functions"
Public Const SHEET_WI As String = "Wi"

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

' ====== Символы ======
Public Const CH_LPAREN As String = "("
Public Const CH_RPAREN As String = ")"
Public Const CH_AND As String = "*"
Public Const CH_OR As String = "+"

' ====== Параметры ======
Public Const R_MAX As Long = 4
Public Const WI_COL_R As String = "A"
Public Const WI_COL_MAX As String = "O"
