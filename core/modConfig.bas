Attribute VB_Name = "modConfig"
' modConfig.bas - Globális beállítások és konstansok

Option Explicit

' Fájl elérési útvonalak
Public Const PART_DETAILS_PATH As String = "P:\_Departments\PW\10. Transfer\Sárdy Péter\excel\pw-plan\part-details.xlsm"
Public Const DEFAULT_SAVE_DIR As String = "\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Documents\pw-plan\"

' Színkódok
Public Const COLOR_HEADER_BG As Long = RGB(26, 26, 26)
Public Const COLOR_HEADER_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_ALTERNATING_1 As Long = RGB(255, 255, 255)
Public Const COLOR_ALTERNATING_2 As Long = RGB(220, 220, 220)
Public Const COLOR_SPECIAL_ROW As Long = RGB(26, 26, 26)
Public Const COLOR_SPECIAL_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_BORDER As Long = RGB(191, 191, 191)

' Táblázat beállítások
Public Const TABLE_NAME As String = "PwPlan"
Public Const TABLE_STYLE As String = "TableStyleMedium9"

' Elrejtendõ oszlopok
Public ColumnsToHide() As String
