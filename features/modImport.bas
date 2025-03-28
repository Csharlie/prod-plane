Public Sub modImport()
    Dim FSO As Object
    Dim folderPath As String

    ' Alapértelmezett importálási mappa
    folderPath = "C:\Users\" & Environ$("USERNAME") & "\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Projektek\vba\projects\prod-plan\"

    ' Fájlrendszer objektum létrehozása
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FolderExists(folderPath) Then
        MsgBox "A megadott mappa nem létezik: " & folderPath, vbExclamation
        Exit Sub
    End If

    ' VBA hozzáférési jogosultság ellenőrzése
    On Error Resume Next
    If ThisWorkbook.VBProject Is Nothing Then
        MsgBox "A VBA projekthez való hozzáférés nincs engedélyezve!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Modulok bejárása és importálása
    ImportModulesFromFolder FSO, folderPath

    Set FSO = Nothing
    MsgBox "Modulok importálása befejeződött!", vbInformation
End Sub

Private Sub ImportModulesFromFolder(FSO As Object, folderPath As String)
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim component As Object
    Dim fileName As String
    Dim fileContent As String

    Set folder = FSO.GetFolder(folderPath)

    ' Fájlok bejárása a mappában
    For Each file In folder.Files
        If LCase(FSO.GetExtensionName(file.Name)) = "bas" Then
            fileName = file.Path
            For Each component In ThisWorkbook.VBProject.VBComponents
                If component.Type = 1 And component.Name & ".bas" = file.Name Then
                    ' Modul előző tartalmának törlése
                    If component.CodeModule.CountOfLines > 0 Then
                        component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
                    End If

                    ' Fájl beolvasása és kód importálása
                    fileContent = ReadFileAsUTF8(fileName)
                    component.CodeModule.AddFromString fileContent
                End If
            Next component
        End If
    Next file

    ' Almappák bejárása
    For Each subFolder In folder.SubFolders
        ImportModulesFromFolder FSO, subFolder.Path
    Next subFolder
End Sub

Private Function ReadFileAsUTF8(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadFileAsUTF8 = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function