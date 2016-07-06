Attribute VB_Name = "Globali"
Option Explicit
Public Enum TipoAzione
   Vuoto
   Nuovo
   Modifica
End Enum



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

