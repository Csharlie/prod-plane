' prodPlan.bas - Központi modul, amely összehangolja a többi modult
' Ez a modul felelős a fő munkafolyamat végrehajtásáért és koordinálásáért

Option Explicit

' Fő eljárás, amely a PW tervezési folyamatot indítja
Public Sub RunPwPlan()
    On Error GoTo ErrorHandler
    
    ' Naplózás indítása
    modLogging.LogActivity "PW Plan folyamat indítása"
    
    ' Adatok betöltése és feldolgozása
    modDataLoader.ProcessData
    
    ' Formázás és megjelenítés
    modFormatter.FormatWorksheet
    
    ' Számítások végrehajtása
    modCalculations.CalculateSums
    
    ' Exportálás
    modExport.ExportToFile
    
    ' Sikeres befejezés
    modLogging.LogActivity "PW Plan folyamat sikeresen befejeződött"
    Exit Sub
    
ErrorHandler:
    modLogging.LogError Err.Description, "RunPwPlan"
    MsgBox "Hiba történt a végrehajtás során: " & Err.Description, vbCritical, "Hiba"
End Sub

' Eredeti PwPlan eljárás, amely továbbra is elérhető a visszafelé kompatibilitás érdekében
Public Sub PwPlan()
    RunPwPlan
End Sub