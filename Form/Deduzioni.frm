VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form Deduzioni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione deduzioni"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictureUNEP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "Deduzioni.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame fraComandi 
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   3000
      Width           =   9855
      Begin VB.CommandButton cmdPrint 
         Height          =   450
         Left            =   3960
         Picture         =   "Deduzioni.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Stampa Schermata"
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "Ri&cerca Deduzioni"
         Height          =   450
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdAnnulla 
         Caption         =   "Esci"
         Height          =   450
         Left            =   7800
         TabIndex        =   25
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   450
         Left            =   5880
         TabIndex        =   24
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.Frame fraMain 
      Height          =   2355
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   9855
      Begin VB.CheckBox ChkAnnullo 
         Caption         =   "Check1"
         DataField       =   "Annullo"
         Height          =   240
         Left            =   1680
         TabIndex        =   26
         Tag             =   "PULISCI"
         Top             =   1920
         Width           =   240
      End
      Begin VB.TextBox txtSigla 
         DataField       =   "SIGLA"
         Height          =   285
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "necessario Sigla Inserimento"
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   19
         Top             =   120
         Width           =   1575
         Begin VB.Label LblAtto 
            Caption         =   "Numero atto : "
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.TextBox TxtNota 
         DataField       =   "Nota"
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   6
         Top             =   600
         Width           =   7725
      End
      Begin TDBDate6Ctl.TDBDate txtDataReg 
         DataField       =   "DataEvasionePratica"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "Deduzioni.frx":058C
         Caption         =   "Deduzioni.frx":06A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Deduzioni.frx":0710
         Keys            =   "Deduzioni.frx":072E
         Spin            =   "Deduzioni.frx":078C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   2.12482833205922E-314
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber txtCompetenze 
         DataField       =   "Importo"
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Deduzioni.frx":07B4
         Caption         =   "Deduzioni.frx":07D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Deduzioni.frx":0840
         Keys            =   "Deduzioni.frx":085E
         Spin            =   "Deduzioni.frx":08A8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "#,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label LblAvvNotAnn 
         Caption         =   "ANNULLATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2955
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label LblAnnullo 
         Caption         =   "Annulla deduzione:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "Sigla : "
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   120
         Width           =   630
      End
      Begin VB.Label LblDescrizioneAtto 
         Height          =   240
         Left            =   3120
         TabIndex        =   18
         Top             =   765
         Width           =   1995
      End
      Begin VB.Label LblDescrCodiceAtto 
         Height          =   240
         Left            =   2250
         TabIndex        =   17
         Tag             =   "PULISCI"
         Top             =   1170
         Width           =   2130
      End
      Begin VB.Label LblDescrSpeseVarie 
         Caption         =   "Descrizione : "
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LblCompetenze 
         Caption         =   "Importo : "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label LblDataReg 
         Caption         =   "Data Registrazione : "
         Height          =   255
         Left            =   225
         TabIndex        =   12
         Top             =   195
         Width           =   1455
      End
   End
   Begin VB.Frame fraTop 
      Height          =   570
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      Begin VB.TextBox TxtCodiceAvvocato 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "XXX"
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   285
      End
      Begin VB.Label LblCodiceA 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LblDescrCodAvv 
         Caption         =   "NOME"
         DataField       =   "NOME"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Tag             =   "XXX"
         Top             =   240
         Width           =   5640
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cassetta :"
         Height          =   255
         Left            =   585
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "Deduzioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numOrdinamento As Integer

Dim PassaLoad As Boolean
Dim LocalLoad As Boolean
Public Azione As TipoAzione
Private sWhere As String
Private moFrmRicerca As FrmRicerca
Private m_ID As Long

Implements IAnagraficForm
Implements IForm


Private Sub ChkAnnullo_Click()
 If ChkAnnullo.value = Checked Then
        LblAvvNotAnn.Visible = True
    Else
        LblAvvNotAnn.Visible = False
    End If
End Sub

Private Sub CmdAnnulla_Click()
If CmdSalva.Enabled Then DeLockRecord m_ID, "DEDUZIONI_UNEP"
Unload Me
'If FindForm("frmRicerca") Then
'    Unload FrmRicerca
'End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
'If FindForm("frmRicerca") Then
'    Unload FrmRicerca
'End If
End Sub
Private Sub cmdPrint_Click()
On Error Resume Next
  PrintForm
End Sub

Private Sub CmdRicercaA_Click()
On Error GoTo ErrHandler
    'Ricerca Avvocato
    
    RicercaPerCodice Me, Azione
    txtDataReg = Date
   
     Exit Sub
ErrHandler:
   If err.Number = SearchErrors.FreeBox Or err.Number = SearchErrors.BrokenBox Or err.Number = SearchErrors.UnknownBox Then
      'TODO
   End If
End Sub
Private Sub CmdRicercaAnag_Click()
    
    Set FrmRicerca.frmCaller = Me
    FrmRicerca.tipo = "Anagrafica"
    FrmRicerca.Filtro = " AND STAT<>'A' And CASSETTAROTTA<>'S'"
    
   
    FrmRicerca.Filtro = FrmRicerca.Filtro & " AND NOT (CODAVV LIKE '525%' OR CODAVV LIKE '393%')"
   
    If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If

    Load FrmRicerca

End Sub

Private Sub CmdRicerca_Click()
    
   
    Set moFrmRicerca = New FrmRicerca
    Set moFrmRicerca.frmCaller = Me
    moFrmRicerca.Titolo = "Ricerca Deduzioni"
    moFrmRicerca.tipo = "Ricerca"
    moFrmRicerca.Filtro = ""
    moFrmRicerca.DefaultOrder = "Order By DataEvasionePratica DESC, NumOrdinamento"
    moFrmRicerca.NCol = 6
    moFrmRicerca.PosizioneCodice = 9
    moFrmRicerca.Tabella = "DEDUZIONI_UNEP"
    moFrmRicerca.isUnep = True
    
    
    moFrmRicerca.Query = "SELECT '' AS Ev, CODAVV AS [Codice], " & _
                "Format(Mid(DataEvasionePratica,7,2) & '/' & Mid(DataEvasionePratica,5,2) & '/' & Mid(DataEvasionePratica,1,4),'dd/mm/yyyy') As [Data Registrazione], " & _
                "Nota, Importo,SIGLA as [Sigla Inserimento], Annullo, NumOrdinamento,IdCod " & _
                "FROM DEDUZIONI_UNEP "
                
   
    moFrmRicerca.Show
    'CmbTribunale.Enabled = False
End Sub

Private Sub CmdSalva_Click()
On Error GoTo ErroreSalvataggio
Dim Msg_Errore  As String
Dim Orario As String
Dim saved As Boolean
  If IsTableLocked("DEDUZIONI_UNEP") Then
       MsgBox "La tabelle Deduzioni è bloccata da un altro utente. Riprovare...", vbInformation
  Else
  Dim errMsg As String
        If TxtNota.Text = "" Then
          errMsg = "La descrizione è obbligatoria"
        End If
       
        
        If errMsg <> "" Then
          MsgBox errMsg, vbExclamation + vbOKOnly, "ATAP"
          Else
          
             SaveSetting "ATAP", "Config", "Sigla", txtSigla.Text
             SaveSetting "ATAP", "Config", "DescrizioneDeduzione", TxtNota.Text
            saved = SalvaTutto(Me, "DEDUZIONI_UNEP", sWhere, False, True)
        
            If Not moFrmRicerca Is Nothing Then
                moFrmRicerca.AggiornaGriglia
            End If
        
            If saved Then DeLockRecord m_ID, "DEDUZIONI_UNEP"
         End If

  End If

Exit Sub

ErroreSalvataggio:

    If CmdSalva.Caption = "&Modifica" Then
        Msg_Errore = "Errore durante la modifica di una Deduzione "
    Else
        Msg_Errore = "Errore durante il salavataggio di una Deduzione "
    End If
    Msg_Errore = Msg_Errore & " - numero: " & err & " - riga: " & Erl & " - messaggio: " & Error(err)
    Orario = (Date & " " & Time)
    ErrLogFile "ErroriAtap.txt", Msg_Errore
    


End Sub


Private Sub Form_Load()
   Me.Move 0, 0
    Azione = TipoAzione.Vuoto
    Call TipoMaschera(Me, Azione)
    txtDataReg.MaxDate = Now + 30
  
    PictureUNEP.Visible = True
   
    TxtNota.Text = GetSetting("ATAP", "Config", "DescrizioneDeduzione", "")
   
    
    
End Sub

Private Property Let IForm_IsLoading(RHS As Boolean)
 PassaLoad = RHS
End Property
Private Property Get IForm_IsLoading() As Boolean
IForm_IsLoading = PassaLoad
End Property
Private Property Let IForm_Where(RHS As String)
 sWhere = RHS
End Property
Private Sub IForm_SetFocus()
 Me.SetFocus
End Sub
Private Sub IForm_RisRicerca()
 Dim SQL As String
Dim rs As ADODB.Recordset
Set rs = newAdoRs

PassaLoad = True
LocalLoad = True
SQL = "SELECT IdCod, DEDUZIONI_UNEP.CODAVV, " & _
      "( Mid(DataEvasionePratica,7,2) & '/' & Mid(DataEvasionePratica,5,2)& '/' & Mid(DataEvasionePratica,1,4)) As DataEvasionePratica, " & _
      "AnagraficaAvvocati.NOME, AnagraficaAvvocati.NumOrdinamento, " & _
      "Importo,Nota, Annullo, " & _
      "SIGLA " & _
      " FROM (DEDUZIONI_UNEP INNER JOIN AnagraficaAvvocati ON DEDUZIONI_UNEP.CODAVV = AnagraficaAvvocati.CODAVV)  " & _
      "WHERE " & sWhere
      
rs.Open SQL, g_Settings.DBConnection
m_ID = -1
If Not rs.EOF Then
   
   Caricacampi Me, rs
   
   Azione = TipoAzione.Modifica
   Call TipoMaschera(Me, Azione)
      m_ID = rs("IDCod")
   If IsRecordLocked("IDCod=" & m_ID, "DEDUZIONI_UNEP") Then
      CmdSalva.Enabled = False
     Else
      CmdSalva.Enabled = True
      LockRecord m_ID, "DEDUZIONI_UNEP"
   End If
 Else
    MsgBox "Il caricamento non è andato a buon fine:" & vbCrLf & "potrebbe non essere presente la Cassetta o il Tribunale corrispondente", vbCritical, "Attenzione"
End If
PassaLoad = False
LocalLoad = True

End Sub






Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocato.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(CsCodAvv As String, oAzione As TipoAzione)
Dim rs As ADODB.Recordset
    Azione = TipoAzione.Nuovo
    Set rs = GetADORecordset("AnagraficaAvvocati", "CodAvv,Nome,numOrdinamento", "CodAvv='" & CsCodAvv & "'", g_Settings.DBConnection)
      m_ID = -1
      txtSigla = GetSetting("ATAP", "Config", "Sigla", "")
    If Not rs.EOF Then
     Call RiempiTestata(Me, rs)
     Call TipoMaschera(Me, Azione)
        
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
