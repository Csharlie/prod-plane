VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDataProcessor.cls - Adatfeldolgozó osztály

Option Explicit

Private pTable As ListObject
Private pOrderColumn As Long
Private pMaterialColumn As Long

Public Property Get Table() As ListObject
    Set Table = pTable
End Property

Public Property Set Table(value As ListObject)
    Set pTable = value
End Property

' További tulajdonságok és metódusok...
