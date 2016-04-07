Attribute VB_Name = "Variabili"
Option Explicit
Global g_Settings As CSettings

Global FinestraAttiva(200) As Variant
Global FormRicerca As String

'dichiaro una costante enumExitMode, che valorizzo quando esco dal form
'path db, per sapere se sono uscito cliccando su Ok o su annulla
Public Enum enumExitMode
    exitCANCEL = 0 'Esco dal form di impostazione del path database con il tasto annulla
    exitOk = 1  'Esco dal form di impostazione del path database con il tasto ok
End Enum


