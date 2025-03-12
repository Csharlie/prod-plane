VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorkbookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsWorkbookManager.cls - Munkafüzet kezelõ osztály

Option Explicit

Private pWorkbook As Workbook
Private pWorksheet As Worksheet

Public Property Get Workbook() As Workbook
    Set Workbook = pWorkbook
End Property

Public Property Set Workbook(value As Workbook)
    Set pWorkbook = value
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

Public Sub Initialize(wb As Workbook, ws As Worksheet)
    Set pWorkbook = wb
    Set pWorksheet = ws
End Sub

' További metódusok...
