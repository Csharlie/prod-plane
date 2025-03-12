Attribute VB_Name = "modConfig"
' modConfig.bas - Glob�lis be�ll�t�sok �s konstansok

Option Explicit

' F�jl el�r�si �tvonalak
Public Const PART_DETAILS_PATH As String = "P:\_Departments\PW\10. Transfer\S�rdy P�ter\excel\pw-plan\part-details.xlsm"
Public Const DEFAULT_SAVE_DIR As String = "\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Documents\pw-plan\"

' Sz�nk�dok
Public Const COLOR_HEADER_BG As Long = RGB(26, 26, 26)
Public Const COLOR_HEADER_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_ALTERNATING_1 As Long = RGB(255, 255, 255)
Public Const COLOR_ALTERNATING_2 As Long = RGB(220, 220, 220)
Public Const COLOR_SPECIAL_ROW As Long = RGB(26, 26, 26)
Public Const COLOR_SPECIAL_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_BORDER As Long = RGB(191, 191, 191)

' T�bl�zat be�ll�t�sok
Public Const TABLE_NAME As String = "PwPlan"
Public Const TABLE_STYLE As String = "TableStyleMedium9"

' Elrejtend� oszlopok
Public ColumnsToHide() As String
