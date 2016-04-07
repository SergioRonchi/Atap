Attribute VB_Name = "CodeProgress"
Option Explicit
Public Sub CloseProgress()

' close the progress form
    
    Unload Progress

End Sub

Public Sub OpenProgress(Testo)

' open the progress form
' display a message in the progress form
' BarVisible : True/false : progress bar visible ?
On Error GoTo OpenProgressError

    Progress.Show
    Progress.Caption = Testo
    Progress.Barra.Visible = True
    Progress.BarraScorr.Visible = True
    Progress.BarraScorr.Width = 0
    Progress.Commento.Visible = True
    Progress.Percentuale.Visible = True
    Progress.Commento.Caption = Trim(Testo)

Exit Sub

OpenProgressError:
    Resume Next

End Sub

Public Sub UpdateProgress(Percent, Optional msg As String)

'update the progress bar with percent (1->100)

Dim valore As Variant

On Error GoTo UpdateProgressError

Static PosOld
        
    If Percent > 0 And Percent <= 100 And Percent >= PosOld + 1 Then
        
        If Percent > 2 Then
            Progress.BarraScorr.Width = (Progress.Barra.Width * Percent / 100) - 50
            DoEvents
        End If
        
        Progress.Percentuale.Caption = CInt(Percent) & " %"
    End If
    If msg <> "" Then Progress.Commento = msg
    PosOld = Int(Percent)
    
    'Progress.Refresh
    

Exit Sub

UpdateProgressError:
 
    Resume Next

End Sub



