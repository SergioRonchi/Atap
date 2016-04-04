VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_DaData As TDBDate
Attribute m_DaData.VB_VarHelpID = -1
Private WithEvents m_AData As TDBDate
Attribute m_AData.VB_VarHelpID = -1
Private WithEvents m_Avvocato As TextBox
Attribute m_Avvocato.VB_VarHelpID = -1

Private WithEvents m_CmdSeleziona As CommandButton
Attribute m_CmdSeleziona.VB_VarHelpID = -1
Private WithEvents m_CmdAnagrafica As CommandButton
Attribute m_CmdAnagrafica.VB_VarHelpID = -1
Private WithEvents m_lblDescription As Label
Attribute m_lblDescription.VB_VarHelpID = -1

Private WithEvents m_searcForm As FrmRicerca
Attribute m_searcForm.VB_VarHelpID = -1
Public Event Validate(IsValid As Boolean)

Public Sub Initialize(Da As TDBDate, A As TDBDate, avvocato As TextBox, _
                      cmdSeleziona As CommandButton, cmdAnagrafica As CommandButton, lbl As Label)
                      
Set m_DaData = Da
Set m_AData = A
Set m_Avvocato = avvocato
Set m_CmdAnagrafica = cmdAnagrafica
Set m_CmdSeleziona = cmdSeleziona
Set m_lblDescription = lbl

Set m_searcForm = New FrmRicerca
m_searcForm.Location = m_CmdAnagrafica.Left + m_CmdAnagrafica.Width + 500
 
 m_AData = ""
 m_DaData = #1/1/1999#

                      
End Sub


Private Sub m_Avvocato_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then m_CmdSeleziona_Click
End Sub

Private Sub m_CmdAnagrafica_Click()
 
    Set m_searcForm.frmCaller = Nothing
    m_searcForm.tipo = "Anagrafica"
    m_searcForm.Filtro = " AND STAT<>'A' And CASSETTAROTTA<>'S'"
    If FindForm("frmRicerca") Then
          Unload m_searcForm
    End If

    Load m_searcForm
End Sub

Private Sub m_CmdSeleziona_Click()
On Error GoTo ErrHandler
   RaiseEvent Validate(True)
   If m_Avvocato = "" Then
      m_lblDescription.Caption = "TUTTE LE CASSETTE"
     
   Else
     RicercaPerCodice

     
   End If
   Exit Sub
ErrHandler:
   Dim msg As String
   If err.Number = SearchErrors.FreeBox Or err.Number = SearchErrors.BrokenBox Or err.Number = SearchErrors.UnknownBox Then
            Select Case err.Number
             Case SearchErrors.BrokenBox
                 msg = "Cassetta Vuota !"
             Case SearchErrors.FreeBox
                 msg = "Cassetta Vuota !"
             Case SearchErrors.UnknownBox
                 msg = "Cassetta libera !"
            End Select
            MsgBox msg, vbInformation, "Attenzione"
            m_Avvocato.SetFocus
            SendKeys "{Home}+{End}"
            RaiseEvent Validate(False)
   End If
   
End Sub
Private Sub RicercaPerCodice()

  
    Dim rs As ADODB.Recordset
    Set rs = GetADORecordset("AnagraficaAvvocati", "CODAVV,STAT,CASSETTAROTTA,Nome,numOrdinamento", "CODAVV='" & m_Avvocato.Text & "'", g_Settings.DBConnection)
    If rs Is Nothing Then
      err.Raise SearchErrors.UnknownBox
    Else
      If rs("CASSETTAROTTA") = "S" Then
          err.Raise SearchErrors.BrokenBox

      ElseIf rs("STAT") = "A" Then
           err.Raise SearchErrors.FreeBox
      Else
        m_lblDescription.Caption = rs("Nome")
        
      End If
    End If

End Sub


Private Sub m_searcForm_AvvocatoSelezionato(codice As String)
       m_Avvocato.Text = codice
       m_CmdSeleziona_Click

End Sub
