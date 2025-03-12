Attribute VB_Name = "prodPlan"
' prodPlan.bas - K�zponti modul, amely �sszehangolja a t�bbi modult

Option Explicit

Public Sub ProductionPlan()
    ' Napl�z�s ind�t�sa
    modLogging.StartLogging
    
    ' Adatok bet�lt�se
    modDataLoader.LoadData
    
    ' Form�z�s �s feldolgoz�s
    modFormatter.FormatWorksheet
    
    ' Sz�m�t�sok v�grehajt�sa
    modCalculations.CalculateSums
    
    ' Export�l�s
    modExport.ExportToFile
    
    ' Napl�z�s befejez�se
    modLogging.EndLogging
    
    MsgBox "A production plan feldolgoz�sa sikeresen befejez�d�tt.", vbInformation
End Sub
