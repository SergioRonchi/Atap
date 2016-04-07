Attribute VB_Name = "Globali"
Option Explicit
Public Enum TipoAzione
   Vuoto
   Nuovo
   Modifica
End Enum

Public Sub ZipFile(zipPath As String, zipName As String, originalPath As String, originalFile As String)
Dim sEXE As String, sPath1 As String, sPath2 As String, sDir As String, r As Integer
sEXE = """" & app.Path & "\zip.exe"""
sPath1 = """" & zipPath & "\" & zipName & ".zip"""
sPath2 = """" & originalPath & "\" & originalFile & """"

sPath1 = Replace(sPath1, "\\", "\")
sPath2 = Replace(sPath2, "\\", "\")
sEXE = Replace(sEXE, "\\", "\")

If Dir(Replace(zipPath & "\" & zipName & ".zip", "\\", "\")) <> "" Then
                     r = MsgBox("Il file esiste già,si vuole sovrascivere?", vbYesNo + vbQuestion)
                     If r = vbYes Then
                        Call ShellAndWait(sEXE & " -j " & sPath1 & " " & sPath2, vbHide)
                        MsgBox "Backup eseguito in " & zipPath, vbOKOnly + vbInformation
                     End If
                     Else
                      Call ShellAndWait(sEXE & " -j " & sPath1 & " " & sPath2, vbHide)
                      MsgBox "Backup eseguito in " & zipPath, vbOKOnly + vbInformation
End If
End Sub

Public Function CompactDatabase(Database As String, dest As String) As Boolean

' Assicurarsi che non esista già un file con

' il nome del database compattato.
On Error GoTo CompEr


' Questa istruzione crea una versione compatta

DBEngine.CompactDatabase Database, dest
CompactDatabase = True
Exit Function

CompEr:
MsgBox err.Description
CompactDatabase = False
End Function

