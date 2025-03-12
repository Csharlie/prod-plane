Attribute VB_Name = "prodPlan"
' prodPlan.bas - Központi modul, amely összehangolja a többi modult

Option Explicit

Public Sub ProductionPlan()
    ' Naplózás indítása
    modLogging.StartLogging
    
    ' Adatok betöltése
    modDataLoader.LoadData
    
    ' Formázás és feldolgozás
    modFormatter.FormatWorksheet
    
    ' Számítások végrehajtása
    modCalculations.CalculateSums
    
    ' Exportálás
    modExport.ExportToFile
    
    ' Naplózás befejezése
    modLogging.EndLogging
    
    MsgBox "A production plan feldolgozása sikeresen befejezõdött.", vbInformation
End Sub
