Attribute VB_Name = "Avvio"

Public Sub Main()

'Modulo di partenza del programma

    'Call LeggeRegistroPerDB(cNomeInRegistry) 'chiama il modulo di registro (registro.bas)
    gDbName = app.Path & "\Storico\Atap.mdb"
    Call ImpostaDatabase '***********
    
    '************************************************************************
    ' Per il momento assegno un valore fisso alla var gPathReport
    ' occorrerà poi creare una procedura(vedi frmPathDB)
    ' per la gestione dei registry
    'gPathReport = "d:\visualbasic\atap\Report"
    gPathReport = ".\Report"
    '************************************************************************
    
    gDbPath = Left(gDbName, Len(gDbName) - 8)
    
    
    Atap.Show
    
    
    
End Sub

Public Sub ImpostaDatabase()
Dim r As Long
On Error GoTo CambiaPathDB 'Attiva la routine di verifica del percorso dove trovare il DB
Inizio:
    Set gWs = Workspaces(0) 'definisco un'area di lavoro entro la quale agisco sul DB; tale area non è permanente
    Set gDb = gWs.OpenDatabase(gDbName) 'assegno l'area di lavoro al DBAcc
Exit Sub
CambiaPathDB:
    r = MsgBox("ATTENZIONE: Devi Porre il database Atap.mdb nella cartella" & vbCrLf & _
           "c:\programmi\Atap\Storico\" & vbCrLf & _
           "Si ricorda di togliere l'attributo di sola lettura al file", vbOKCancel + vbInformation, "ATAP")
    err.Clear
    If r = vbOK Then GoTo Inizio Else End
           
    frmPathDB.Show 1  'visualizzo con proprietà modale la frmPathDB
    If frmPathDB.ExitMode = exitCANCEL Then 'se esco premendo il tasto Annulla, allora, chiudo la form
        End
    Else
        Resume 'Riprende l'esecuzione alla stessa riga che ha generato l'errore.
    End If
End Sub
