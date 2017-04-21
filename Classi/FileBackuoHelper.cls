VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileBackuoHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub BackUp(Path As String)
Dim sEXE As String, sPath1 As String, sPath2 As String, sDir As String, r As Integer
Dim fName As String

On Error GoTo errback

 Screen.MousePointer = vbHourglass
 SafeKill g_Settings.AtapUserTempFolder & "\*.mdb"
 Dim dbFile As String
 Dim zippedFile As String
 Dim compactedFile As String
 Dim temporaryCopiedFile As String
 

 dbFile = GetFile(g_Settings.dbFile)
 zippedFile = GetFileNameWithoutExtension(g_Settings.dbFile) & "_" & GetTimeStamp
 compactedFile = g_Settings.AtapUserTempFolder & "\" & dbFile
 g_Settings.DBConnection.Close
 
 DBEngine.CompactDatabase g_Settings.dbFile, compactedFile
 
 
 If Dir(compactedFile) <> "" Then
        zipFile Path, zippedFile, g_Settings.AtapUserTempFolder, dbFile
 End If
 Screen.MousePointer = vbDefault
 
 
g_Settings.ConnettiDB g_Settings.dbFile

 Exit Sub
 
errback:
Screen.MousePointer = vbDefault
g_Settings.ConnettiDB g_Settings.dbFile
 MsgBox "Creazione fallita del backup in " & vbCrLf & Path & vbCrLf & err.Description, vbOKOnly + vbExclamation, "Atap"

End Sub

Public Sub zipFile(zipPath As String, zipName As String, originalPath As String, originalFile As String, Optional msg)
Dim sEXE As String, sPath1 As String, sPath2 As String, sDir As String, r As Integer
sEXE = """" & app.Path & "\zip.exe"""
sPath1 = """" & zipPath & "\" & zipName & ".zip"""
sPath2 = """" & originalPath & "\" & originalFile & """"

sPath1 = Replace(sPath1, "\\", "\")
sPath2 = Replace(sPath2, "\\", "\")
sEXE = Replace(sEXE, "\\", "\")

If Dir(Replace(zipPath & "\" & zipName & ".zip", "\\", "\")) <> "" Then
                     r = MsgBox("Il file esiste già. Sovrascrivere?", vbYesNo + vbQuestion)
                     If r = vbYes Then
                        Call ShellAndWait(sEXE & " -j " & sPath1 & " " & sPath2, vbHide)
                        Screen.MousePointer = vbDefault
                        MsgBox "Eseguito backup in " & " " & zipPath, vbOKOnly + vbInformation
                     End If
                     Else
                      Call ShellAndWait(sEXE & " -j " & sPath1 & " " & sPath2, vbHide)
                      Screen.MousePointer = vbDefault
                      If msg <> "" Then
                          MsgBox msg & " " & zipPath, vbOKOnly + vbInformation
                        Else
                         MsgBox "Eseguito  backup in " & " " & zipPath, vbOKOnly + vbInformation
                      End If
                     
End If
End Sub


