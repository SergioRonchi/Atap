Attribute VB_Name = "Liquidazione"
Option Explicit
Dim TrasferimentoOK As Boolean


Public Sub SvuotaTabellaSaldi(CodiceAvv As String)
Dim SQL As String
SQL = "UPDATE Saldi SET Saldi.SaldoAdemp = 0, Saldi.SaldoAdempEuro = 0, " & _
    "Saldi.SaldoSfpg = 0, Saldi.SaldoSfpgEuro = 0, Saldi.SaldoNotif = 0, " & _
    "Saldi.SaldoNotifEuro = 0, Saldi.SaldoDecrIng = 0, " & _
    "Saldi.SaldoDecrIngEuro = 0, Saldi.SaldoTotale = 0, " & _
    "Saldi.SaldoTotaleEuro = 0, Saldi.Commento = '', " & _
    "Saldi.Stato = 'N', Saldi.Chiusura = '" & Format(Now, "yyyymmdd") & "' " & _
    "WHERE Saldi.Codice='" & CodiceAvv & "';"
g_Settings.DBConnection.Execute SQL

End Sub
Public Sub SvuotaTabellaSaldiUNEP(CodiceAvv As String)
Dim SQL As String
SQL = "UPDATE SaldiUNEP SET SaldiUNEP.SaldoAdemp = 0, SaldiUNEP.SaldoAdempEuro = 0, " & _
    "SaldiUNEP.SaldoSfpg = 0, SaldiUNEP.SaldoSfpgEuro = 0, SaldiUNEP.SaldoNotif = 0, " & _
    "SaldiUNEP.SaldoNotifEuro = 0, SaldiUNEP.SaldoDecrIng = 0, " & _
    "SaldiUNEP.SaldoDecrIngEuro = 0, SaldiUNEP.SaldoTotale = 0, " & _
    "SaldiUNEP.SaldoTotaleEuro = 0, SaldiUNEP.Commento = '', " & _
    "SaldiUNEP.Stato = 'N', SaldiUNEP.Chiusura = '" & Format(Now, "yyyymmdd") & "' " & _
    "WHERE SaldiUNEP.Codice='" & CodiceAvv & "';"
g_Settings.DBConnection.Execute SQL

End Sub
Public Sub LiberaCasella(CodiceAvv As String)
Dim SQL As String
SQL = "UPDATE AnagraficaAvvocati SET AnagraficaAvvocati.NOME = '', " & _
      "AnagraficaAvvocati.INDIRI = '', AnagraficaAvvocati.LOCALI = '', " & _
      "AnagraficaAvvocati.PROV = '', AnagraficaAvvocati.CAP = '', " & _
      "AnagraficaAvvocati.TELEFCELL = '', AnagraficaAvvocati.TELEF = '', " & _
      "AnagraficaAvvocati.EMAIL = '', AnagraficaAvvocati.FAX = '', " & _
      "AnagraficaAvvocati.PEC = '', AnagraficaAvvocati.MAIL2 = '', " & _
      "AnagraficaAvvocati.PIVA = '', AnagraficaAvvocati.CFISC = '', " & _
      "AnagraficaAvvocati.NOTE1 = '', AnagraficaAvvocati.NOTE2 = '', " & _
      "AnagraficaAvvocati.NOTE3 = '', AnagraficaAvvocati.STAT = 'A', " & _
      "AnagraficaAvvocati.FILLER = '', AnagraficaAvvocati.AFAT = 'N', " & _
      "AnagraficaAvvocati.SALDO = 0, AnagraficaAvvocati.CassettaRotta = 'N' " & _
      "WHERE AnagraficaAvvocati.CodAvv='" & CodiceAvv & "';"
g_Settings.DBConnection.Execute SQL
g_Settings.DBConnection.Execute "DELETE * FROM USUFRUENTI WHERE CodAvv='" & CodiceAvv & "';"
End Sub

