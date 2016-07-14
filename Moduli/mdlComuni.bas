Attribute VB_Name = "mdlComuni"
Option Explicit


Public Property Get K_TUTTI() As String
 K_TUTTI = " -- Tutte le cassette -- "
End Property
Public Sub RicercaPerCodice(frm As IAnagraficForm, Azione As TipoAzione)
'Ricerca Avvocato
Dim txtDataReg As String
With frm
    txtDataReg = Date
    Dim rs As ADODB.Recordset
    Set rs = GetADORecordset("AnagraficaAvvocati", "CODAVV,STAT,CASSETTAROTTA", "CODAVV='" & .GetCodiceAvvocato() & "'", g_Settings.DBConnection)
    If rs Is Nothing Then
      MsgBox "La cassetta non esiste!", vbInformation, "Attenzione"
      .SelectCodiceAvvocato
'      .TxtCodiceAvvocato.SetFocus
'      SendKeys "{Home}+{End}"
      err.Raise SearchErrors.UnknownBox

    Else
      If rs("CASSETTAROTTA") = "S" Then
          MsgBox "Cassetta Vuota !", vbCritical, "Attenzione"
          .SelectCodiceAvvocato
'          .TxtCodiceAvvocato.SetFocus
'          SendKeys "{Home}+{End}"
          err.Raise SearchErrors.BrokenBox

      ElseIf rs("STAT") = "A" Then
           MsgBox "Cassetta libera !", vbInformation, "Attenzione"
           .SelectCodiceAvvocato
'           .TxtCodiceAvvocato.SetFocus
'           SendKeys "{Home}+{End}"
           err.Raise SearchErrors.FreeBox
      Else
        .RisultatoRicerca .GetCodiceAvvocato(), Azione
        
      End If
    End If
End With
End Sub


Public Sub PulisciTestata(fra As Form)
With fra
    .TxtCodiceAvvocato.Text = ""
    .LblCodiceA.Caption = ""
    .LblDescrCodAvv.Caption = "Descrizione: "
End With
End Sub
Public Sub PulisciCampi(frm As Form)
 Dim c As Control
 For Each c In frm
  If TypeOf c Is TDBDate Then
     c.Text = ""
  ElseIf TypeOf c Is TextBox Then
     c.Text = ""
  ElseIf TypeOf c Is TDBNumber Then
     c.value = 0
  ElseIf TypeOf c Is Label Then
    If c.Tag = "PULISCI" Then c.Caption = ""
  ElseIf TypeOf c Is CheckBox Then
    If c.Tag = "PULISCI" Then c.value = Unchecked
  ElseIf TypeOf c Is ComboBox Then
    c.Text = ""
  End If
 Next
 
 End Sub

