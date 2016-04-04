VERSION 5.00
Begin VB.Form GestioneCambioCassetta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Cambio Cassetta"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmTestataAdemp 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Command1 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   450
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1860
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Enabled         =   0   'False
         Height          =   555
         Left            =   8415
         TabIndex        =   9
         Top             =   1620
         Width           =   1260
      End
      Begin VB.TextBox TxtNewCod 
         Height          =   285
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1350
         Width           =   1485
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   450
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1860
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   2
         Top             =   225
         Width           =   1410
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   3180
         TabIndex        =   1
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "Cassetta nuova"
         Height          =   255
         Left            =   225
         TabIndex        =   8
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   9810
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cassetta Vecchia"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label LblDescrCodAvv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   5235
      End
      Begin VB.Label LblCodiceA 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
End
Attribute VB_Name = "GestioneCambioCassetta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipo As Integer
Implements IAnagraficForm
Private Sub CmdRicercaA_Click()
On Error GoTo ErrHandler
   RicercaPerCodice Me, TipoAzione.Nuovo
   Exit Sub
ErrHandler:
   If err.Number = SearchErrors.FreeBox Or err.Number = SearchErrors.BrokenBox Or err.Number = SearchErrors.UnknownBox Then
      'TODO
   End If
End Sub

Private Sub CmdRicercaAnag_Click()
    tipo = 1
    Set FrmRicerca.frmCaller = Me
    FrmRicerca.tipo = "Anagrafica"
    FrmRicerca.Filtro = " AND STAT<>'A'"
        If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If
   FrmRicerca.Show

End Sub

Private Sub CmdSalva_Click()
    Dim newCod, oldCod As String
    Dim libero As Boolean
    Dim Response As Variant
    Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
        newCod = TxtNewCod.Text
        oldCod = TxtCodiceAvvocato.Text
        If newCod = "" Then
            MsgBox "Nuovo codice Obbligatorio!", vbInformation, "Attenzione"
            TxtNewCod.SetFocus
            Exit Sub
        End If
        libero = isFree(newCod)
        If libero = True Then
            updateAnagraficaAvvocati oldCod, newCod
            
            MsgBox "Salvataggio effettuato !", vbInformation, "Attenzione"
            PulisciCampi
        End If
    End If
End Sub




Private Function isFree(ByVal cod As String) As Boolean
    Dim free As Boolean
    Dim rs As ADODB.Recordset
    cod = UCase(cod)
    Set rs = GetADORecordset("AnagraficaAvvocati", "STAT,cassettaRotta", "CodAvv='" & cod & "'", g_Settings.DBConnection)
    If Not rs Is Nothing Then
         If ControlloNULL(rs!STAT) = "A" And ControlloNULL(rs!cassettaRotta) = "N" Then
            'Cassetta libera
            free = True
         Else
            free = False
            If ControlloNULL(rs!STAT) = "A" And ControlloNULL(rs!cassettaRotta) = "S" Then
                'Cassetta Vuota
                MsgBox "Cassetta Vuota!", vbCritical, "Attenzione"
            End If
            If ControlloNULL(rs!STAT) <> "A" Then
                'Cassetta occupata
                MsgBox "Cassetta occupata!", vbCritical, "Attenzione"
            End If
         End If
    Else
         'Cassetta nuova
         free = True
         'Crea nuova cassetta
         creaNuovaCassetta cod
         'Crea nuova riga saldo x nuova cassetta
         creaNuovaRigaSaldo cod
    End If
    
    isFree = free
End Function

Private Sub creaNuovaCassetta(cod As String)
    g_Settings.DBConnection.Execute "INSERT INTO AnagraficaAvvocati (CodAvv) VALUES ('" & cod & "')"
    
End Sub

Private Sub creaNuovaRigaSaldo(cod As String)
    g_Settings.DBConnection.Execute "INSERT INTO Saldi (Codice) VALUES ('" & cod & "')"
    
End Sub


Private Sub updateAnagraficaAvvocati(ByVal oldCod As String, ByVal newCod As String)
    Dim qry As String
    On Error GoTo FINE
    g_Settings.DBConnection.BeginTrans
    ' Query di aggiornamento tabella AnagraficaAvvocati
    qry = ""
    qry = "UPDATE AnagraficaAvvocati, AnagraficaAvvocati AS AnagraficaAvvocati_Old"
    qry = qry + " SET AnagraficaAvvocati.NumOrdinamento = [AnagraficaAvvocati_Old].[NumOrdinamento],AnagraficaAvvocati.NOME = [AnagraficaAvvocati_Old].[NOME], AnagraficaAvvocati.INDIRI = [AnagraficaAvvocati_Old].[INDIRI],"
    qry = qry + " AnagraficaAvvocati.LOCALI = [AnagraficaAvvocati_Old].[LOCALI], AnagraficaAvvocati.PROV = [AnagraficaAvvocati_Old].[PROV], AnagraficaAvvocati.CAP = [AnagraficaAvvocati_Old].[CAP], "
    qry = qry + " AnagraficaAvvocati.TELEFCELL = [AnagraficaAvvocati_Old].[TELEFCELL], AnagraficaAvvocati.TELEF = [AnagraficaAvvocati_Old].[TELEF], AnagraficaAvvocati.EMAIL = [AnagraficaAvvocati_Old].[EMAIL], "
    qry = qry + " AnagraficaAvvocati.FAX = [AnagraficaAvvocati_Old].[FAX], AnagraficaAvvocati.PIVA = [AnagraficaAvvocati_Old].[PIVA], AnagraficaAvvocati.CFISC = [AnagraficaAvvocati_Old].[CFISC], "
    qry = qry + " AnagraficaAvvocati.NOTE1 = [AnagraficaAvvocati_Old].[NOTE1], AnagraficaAvvocati.NOTE2 = [AnagraficaAvvocati_Old].[NOTE2], AnagraficaAvvocati.NOTE3 = [AnagraficaAvvocati_Old].[NOTE3], AnagraficaAvvocati.STAT = 'V', "
    qry = qry + " AnagraficaAvvocati.CassettaRotta = [AnagraficaAvvocati_Old].[CassettaRotta], AnagraficaAvvocati.AFAT = [AnagraficaAvvocati_Old].[AFAT], AnagraficaAvvocati.SALDO = [AnagraficaAvvocati_Old].[SALDO]"
    qry = qry + " WHERE (([AnagraficaAvvocati].[CODAVV]='" + newCod + "' And [AnagraficaAvvocati_Old].[CODAVV]='" + oldCod + "'))"
    'Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
    'Query di aggiornamento tabella AnagraficaAvvocati * Cassetta Vuota
    qry = "UPDATE AnagraficaAvvocati"
    qry = qry + " SET AnagraficaAvvocati.NOME = '', AnagraficaAvvocati.INDIRI = '',"
    qry = qry + " AnagraficaAvvocati.LOCALI = '', AnagraficaAvvocati.PROV = '', AnagraficaAvvocati.CAP = '', "
    qry = qry + " AnagraficaAvvocati.TELEFCELL = '', AnagraficaAvvocati.TELEF = '', AnagraficaAvvocati.EMAIL = '', "
    qry = qry + " AnagraficaAvvocati.FAX = '', AnagraficaAvvocati.PIVA = '', AnagraficaAvvocati.CFISC = '', "
    qry = qry + " AnagraficaAvvocati.NOTE1 = '', AnagraficaAvvocati.NOTE2 = '', AnagraficaAvvocati.NOTE3 = '', AnagraficaAvvocati.STAT = 'A', "
    qry = qry + " AnagraficaAvvocati.CassettaRotta = 'S'"
    ', AnagraficaAvvocati.AFAT = '', AnagraficaAvvocati.SALDO = ''"
    qry = qry + " WHERE (([AnagraficaAvvocati].[CODAVV]='" + oldCod + "'))"
    ''Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
 qry = ""
    qry = " UPDATE Saldi, Saldi AS Saldi_Old "
    qry = qry + " SET Saldi.Chiusura = [Saldi_Old].[Chiusura], Saldi.SaldoAdemp = [Saldi_Old].[SaldoAdemp], "
    qry = qry + " Saldi.SaldoAdempEuro = [Saldi_Old].[SaldoAdempEuro], Saldi.SaldoSfpg = [Saldi_Old].[SaldoSfpg],"
    qry = qry + " Saldi.SaldoSfpgEuro = [Saldi_Old].[SaldoSfpgEuro], Saldi.SaldoNotif = [Saldi_Old].[SaldoNotif], "
    qry = qry + " Saldi.SaldoNotifEuro = [Saldi_Old].[SaldoNotifEuro], Saldi.SaldoDecrIng = [Saldi_Old].[SaldoDecrIng],"
    qry = qry + " Saldi.SaldoDecrIngEuro = [Saldi_Old].[SaldoDecrIngEuro], Saldi.Stato = [Saldi_Old].[Stato], "
    qry = qry + " Saldi.SaldoTotale = [Saldi_Old].[SaldoTotale], Saldi.SaldoTotaleEuro = [Saldi_Old].[SaldoTotaleEuro], "
    qry = qry + " Saldi.PROG_Saldi = [Saldi_Old].[PROG_Saldi], Saldi.Commento = [Saldi_Old].[Commento], Saldi.NumOrdinamento = [Saldi_Old].[NumOrdinamento]"
    qry = qry + " WHERE (([Saldi].[Codice]='" + newCod + "'  And [Saldi_Old].[Codice]='" + oldCod + "'))"
    ''Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
    qry = ""
    qry = " UPDATE SaldiUNEP, SaldiUNEP AS Saldi_Old "
    qry = qry + " SET SaldiUNEP.Chiusura = [Saldi_Old].[Chiusura], SaldiUNEP.SaldoAdemp = [Saldi_Old].[SaldoAdemp], "
    qry = qry + " SaldiUNEP.SaldoAdempEuro = [Saldi_Old].[SaldoAdempEuro], SaldiUNEP.SaldoSfpg = [Saldi_Old].[SaldoSfpg],"
    qry = qry + " SaldiUNEP.SaldoSfpgEuro = [Saldi_Old].[SaldoSfpgEuro], SaldiUNEP.SaldoNotif = [Saldi_Old].[SaldoNotif], "
    qry = qry + " SaldiUNEP.SaldoNotifEuro = [Saldi_Old].[SaldoNotifEuro], SaldiUNEP.SaldoDecrIng = [Saldi_Old].[SaldoDecrIng],"
    qry = qry + " SaldiUNEP.SaldoDecrIngEuro = [Saldi_Old].[SaldoDecrIngEuro], SaldiUNEP.Stato = [Saldi_Old].[Stato], "
    qry = qry + " SaldiUNEP.SaldoTotale = [Saldi_Old].[SaldoTotale], SaldiUNEP.SaldoTotaleEuro = [Saldi_Old].[SaldoTotaleEuro], "
    qry = qry + " SaldiUNEP.PROG_Saldi = [Saldi_Old].[PROG_Saldi], SaldiUNEP.Commento = [Saldi_Old].[Commento], SaldiUNEP.NumOrdinamento = [Saldi_Old].[NumOrdinamento]"
    qry = qry + " WHERE (([SaldiUNEP].[Codice]='" + newCod + "'  And [Saldi_Old].[Codice]='" + oldCod + "'))"
    ''Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
    ' Query di aggiornamento tabella Saldi * Cassetta Vuota
    qry = ""
    qry = " UPDATE Saldi"
    qry = qry + " SET Saldi.SaldoAdemp = '0', "
    qry = qry + " Saldi.SaldoAdempEuro = '0', Saldi.SaldoSfpg = '0',"
    qry = qry + " Saldi.SaldoSfpgEuro = '0', Saldi.SaldoNotif = '0', "
    qry = qry + " Saldi.SaldoNotifEuro = '0', Saldi.SaldoDecrIng = '0',"
    qry = qry + " Saldi.SaldoDecrIngEuro = '0', Saldi.Stato = '',"
    qry = qry + " Saldi.SaldoTotale = '0', Saldi.SaldoTotaleEuro = '0',"
    qry = qry + " Saldi.PROG_Saldi = '0', Saldi.Commento = ''"
    qry = qry + " WHERE (([Saldi].[Codice]='" + oldCod + "'))"
    ''Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
     qry = ""
    qry = " UPDATE SaldiUNEP"
    qry = qry + " SET SaldiUNEP.SaldoAdemp = '0', "
    qry = qry + " SaldiUNEP.SaldoAdempEuro = '0', Saldi.SaldoSfpg = '0',"
    qry = qry + " SaldiUNEP.SaldoSfpgEuro = '0', Saldi.SaldoNotif = '0', "
    qry = qry + " SaldiUNEP.SaldoNotifEuro = '0', Saldi.SaldoDecrIng = '0',"
    qry = qry + " SaldiUNEP.SaldoDecrIngEuro = '0', Saldi.Stato = '',"
    qry = qry + " SaldiUNEP.SaldoTotale = '0', Saldi.SaldoTotaleEuro = '0',"
    qry = qry + " SaldiUNEP.PROG_Saldi = '0', SaldiUNEP.Commento = ''"
    qry = qry + " WHERE (([SaldiUNEP].[Codice]='" + oldCod + "'))"
    ''Debug.print qry
    g_Settings.DBConnection.Execute (qry)
    
    

qry = "UPDATE ADEMPI SET ADEMPI.CODAVV = '" + newCod + "'"
qry = qry + " WHERE(((ADEMPI.CODAVV)='" + oldCod + "'))"
''Debug.print qry
g_Settings.DBConnection.Execute (qry)

qry = "UPDATE Usufruenti SET Usufruenti.CODAVV = '" + newCod + "'"
qry = qry + " WHERE(((Usufruenti.CODAVV)='" + oldCod + "'))"
''Debug.print qry
g_Settings.DBConnection.Execute (qry)

qry = "UPDATE DecretiIngiuntivi SET DecretiIngiuntivi.CODAVV = '" + newCod + "'"
qry = qry + " WHERE (((DecretiIngiuntivi.CODAVV)='" + oldCod + "'))"
g_Settings.DBConnection.Execute (qry)


qry = "UPDATE SFRATTI SET SFRATTI.CODAVV = '" + newCod + "'"
qry = qry + " WHERE (((SFRATTI.CODAVV)='" + oldCod + "'))"
g_Settings.DBConnection.Execute (qry)

qry = "UPDATE Notifiche SET Notifiche.CODAVV = '" + newCod + "'"
qry = qry + " WHERE (((Notifiche.CODAVV)='" + oldCod + "'))"
g_Settings.DBConnection.Execute (qry)

    g_Settings.DBConnection.CommitTrans
    Exit Sub
FINE:
    MsgBox "Cambio cassetta non riuscito!!", vbCritical + vbOKOnly
    g_Settings.DBConnection.RollbackTrans
End Sub


Public Sub PulisciCampi()
     TxtCodiceAvvocato.Text = ""
     LblDescrCodAvv.Caption = ""
     TxtNewCod.Text = ""
End Sub



Private Sub Command1_Click()
    tipo = 2
    Set FrmRicerca.frmCaller = Me
    FrmRicerca.tipo = "Anagrafica"
    FrmRicerca.Filtro = " AND STAT='A'"
        If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If
   FrmRicerca.Show

End Sub

Private Sub TxtCodiceAvvocato_Change()
 
    CmdSalva.Enabled = TxtCodiceAvvocato.Text <> "" And TxtNewCod <> ""
 
End Sub

Private Sub TxtNewCod_Change()
CmdSalva.Enabled = TxtCodiceAvvocato.Text <> "" And TxtNewCod <> ""
End Sub


Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocato.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(sCodAvv As String, oAzione As TipoAzione)
Dim rs As ADODB.Recordset
    'Nuovo adempimento
    
    Set rs = GetADORecordset("AnagraficaAvvocati", "CodAvv,Nome ", "CodAvv='" & sCodAvv & "'", g_Settings.DBConnection)
    
    If Not rs.EOF Then
     If tipo = 1 Then
       Call RiempiTestata(Me, rs)
      Else
        TxtNewCod = sCodAvv
     End If
     
    Else
        MsgBox "Il caricamento della testata non è andato a buon fine provare a rieseguire l'operazione!", vbCritical, "Attenzione"
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub IAnagraficForm_SelectCodiceAvvocato()
 TxtCodiceAvvocato.SetFocus
 SendKeys "{Home}+{End}"
End Sub


