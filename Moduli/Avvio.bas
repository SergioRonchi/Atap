Attribute VB_Name = "Avvio"
Option Explicit
Public Sub Main()
   
'Modulo di partenza del programma
    If app.PrevInstance = True Then
      End
    End If
    
    Set g_Settings = New CSettings

    Atap.Show

End Sub

