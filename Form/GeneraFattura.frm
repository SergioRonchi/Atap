VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form GeneraFattura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Genera Fattura"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmMetodoStampa 
      Caption         =   "Modalità Stampa"
      Height          =   645
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   4980
      Begin VB.OptionButton OptModSt 
         Caption         =   "Diretta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   270
         Width           =   1680
      End
      Begin VB.OptionButton OptModSt 
         Caption         =   "Anteprima"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1680
      End
   End
   Begin VB.Frame FrmComandi 
      Height          =   780
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   4980
      Begin VB.CommandButton CmdAnnulla 
         Caption         =   "E&sci"
         Height          =   500
         Left            =   3360
         TabIndex        =   25
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   500
         Left            =   120
         TabIndex        =   8
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton CmdPulisci 
         Caption         =   "&Pulisci"
         Height          =   500
         Left            =   1740
         TabIndex        =   9
         Top             =   195
         Width           =   1500
      End
   End
   Begin VB.Frame FrmImpFat 
      Caption         =   " Impostazioni Fattura "
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   4980
      Begin VB.Frame FrmValuta 
         Caption         =   " Valuta "
         Height          =   990
         Left            =   4035
         TabIndex        =   19
         Top             =   150
         Width           =   855
         Begin VB.Label Label1 
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   285
            TabIndex        =   23
            Top             =   270
            Width           =   420
         End
      End
      Begin TDBNumber6Ctl.TDBNumber txtImponibile 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "GeneraFattura.frx":0000
         Caption         =   "GeneraFattura.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "GeneraFattura.frx":008C
         Keys            =   "GeneraFattura.frx":00AA
         Spin            =   "GeneraFattura.frx":00F4
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
      Begin TDBNumber6Ctl.TDBNumber txtNFat 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "GeneraFattura.frx":011C
         Caption         =   "GeneraFattura.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "GeneraFattura.frx":01A8
         Keys            =   "GeneraFattura.frx":01C6
         Spin            =   "GeneraFattura.frx":0210
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#######0;;"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "#######0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2000000
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   159055873
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate txtDataFat 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "GeneraFattura.frx":0238
         Caption         =   "GeneraFattura.frx":0350
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "GeneraFattura.frx":03BC
         Keys            =   "GeneraFattura.frx":03DA
         Spin            =   "GeneraFattura.frx":0438
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
         Format          =   ""
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   36161
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   ""
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   2.12482833205922E-314
         CenturyMode     =   0
      End
      Begin VB.Label lblImponibile 
         Caption         =   "Imponibile :"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1245
         Width           =   960
      End
      Begin VB.Label lblNFat 
         Caption         =   "N° Fattura :"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   825
         Width           =   960
      End
      Begin VB.Label lblDataFat 
         Caption         =   "Data Fattura :"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   990
      End
   End
   Begin VB.Frame FrmImpAnag 
      Caption         =   " Impostazioni  Anagrafiche "
      Height          =   2835
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   4980
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   465
         Left            =   3360
         TabIndex        =   2
         Top             =   285
         Width           =   1410
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   0
         Top             =   330
         Width           =   1395
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2835
         TabIndex        =   1
         Top             =   345
         Width           =   330
      End
      Begin VB.Label lblCAP 
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label lblProv 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   2400
         Width           =   3690
      End
      Begin VB.Label lblLocali 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   2160
         Width           =   3690
      End
      Begin VB.Label lblIndiri 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   1920
         Width           =   3570
      End
      Begin VB.Label LblCodiceA 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblDesPIVA 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   1605
         Width           =   3570
      End
      Begin VB.Label lblPIVA 
         Caption         =   "Partita IVA :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label LblDescrCodAvv 
         ForeColor       =   &H000000C0&
         Height          =   510
         Left            =   1200
         TabIndex        =   13
         Top             =   870
         Width           =   3705
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta :"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label LblDescr 
         Caption         =   "Descrizione :"
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   870
         Width           =   1110
      End
   End
End
Attribute VB_Name = "GeneraFattura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Azione As TipoAzione
Private moFatturaManager As New CFattureManager
Private WithEvents m_Printmanager As Stampa
Attribute m_Printmanager.VB_VarHelpID = -1
Implements IAnagraficForm
Implements IForm



Private Sub CmdAnnulla_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub

Private Sub CmdOK_Click()
On Error GoTo FINE
    If LblDescrCodAvv.Caption = "" Then
        MsgBox "Selezionare un avvocato!", vbInformation, "Attenzione"
        TxtCodiceAvvocato.SetFocus
        Exit Sub
    End If
    If txtNFat.value = 0 Then
        MsgBox "Inserire il numero fattura!", vbInformation, "Attenzione"
        txtNFat.SetFocus
        Exit Sub
    End If
    If txtImponibile.value = 0 Then
        MsgBox "Inserire l'imponibile!", vbInformation, "Attenzione"
        txtImponibile.SetFocus
        Exit Sub
    End If
    Set m_Printmanager = New Stampa
    If Not IsPrtTableLocked("PrtFattProv") Then
        LockPrtTable ("PrtFattProv")
        GeneraFatturaProv
        Call m_Printmanager.gestioneReport("PrtFattProv", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "FatturaProv.rpt", 2)
        moFatturaManager.SaveNumber (txtNFat.value)
        txtNFat.value = moFatturaManager.GetNextNumber()
      Else
       MsgBox "Fatture Bloccate da un altro utente. Riprovare.", vbInformation
    End If
  Exit Sub
FINE:
  MsgBox "Controllare i dati.", vbInformation
End Sub

Private Sub CmdRicercaA_Click()
On Error GoTo ErrHandler
   RicercaPerCodice Me, Azione
   Exit Sub
ErrHandler:
   If err.Number = SearchErrors.FreeBox Or err.Number = SearchErrors.BrokenBox Or err.Number = SearchErrors.UnknownBox Then
      'TODO
   End If
End Sub

Private Sub CmdRicercaAnag_Click()
    Azione = TipoAzione.Nuovo
    Set FrmRicerca.frmCaller = Me
    FrmRicerca.tipo = "Anagrafica"
    FrmRicerca.Filtro = " AND STAT<>'A' And CASSETTAROTTA<>'S' AND AFAT='S'"
        If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If
   FrmRicerca.Show

End Sub

Private Sub Form_Load()
   txtDataFat = Date
   txtNFat.value = moFatturaManager.GetNextNumber()
End Sub


Public Sub GeneraFatturaProv()

   On Error GoTo FINE
    
    Dim qry As String
    g_Settings.DBConnection.BeginTrans
    qry = "DELETE FROM PrtFattProv;"
    g_Settings.DBConnection.Execute qry

    g_Settings.DBConnection.Execute "INSERT INTO PrtFattProv (CodAvv,NOME,PIVA,NumeroFattura,Datafattura,Indiri,Locali,Prov,ImportoIVA,CAP,CompAdemp,Valuta) VALUES (" & _
                     "'" & LblCodiceA & "','" & Replace(LblDescrCodAvv, "'", "''") & "','" & Replace(lblDesPIVA, "'", "''") & "','" & txtNFat & "','" & txtDataFat.Text & _
                     "','" & Replace(lblIndiri, "'", "''") & "','" & Replace(lblLocali, "'", "''") & "','" & Replace(lblProv, "'", "''") & "','" & g_Settings.IVA & "','" & Replace(lblCAP, "'", "''") & "','" & txtImponibile & "','Euro')"
    g_Settings.DBConnection.CommitTrans
    Exit Sub
FINE:
   MsgBox err.Description, vbCritical

   g_Settings.DBConnection.RollbackTrans
    
End Sub

Private Property Let IForm_IsLoading(RHS As Boolean)

End Property

Private Property Get IForm_IsLoading() As Boolean

End Property

Private Sub IForm_RisRicerca()

End Sub

Private Sub IForm_SetFocus()
 Me.SetFocus
End Sub

Private Property Let IForm_Where(RHS As String)

End Property

Private Sub m_Printmanager_StampaEseguita(table As String)
 DelockPrtTable ("PrtFattProv")
End Sub

Private Sub TxtCodiceAvvocato_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdRicercaA_Click
End Sub

Private Sub txtImponibile_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocato.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(sCodAvv As String, oAzione As TipoAzione)

  Dim rs As ADODB.Recordset
    'Nuovo adempimento
    
    Set rs = GetADORecordset("AnagraficaAvvocati", "CodAvv,Nome,numOrdinamento,Locali,Indiri,prov,PIVA,CAP, AFAT", "CodAvv='" & sCodAvv & "'", g_Settings.DBConnection)
    CmdOK.Enabled = False
    If Not rs.EOF Then
     
            Call RiempiTestata(Me, rs)
            lblIndiri = rs("Indiri")
            lblLocali = rs("LocalI")
            lblProv = rs("Prov")
            lblDesPIVA = rs("PIVA")
            lblCAP = rs("CAP")
            
        If rs("AFAT") = "S" And rs("PIVA") <> "" Then
            CmdOK.Enabled = True
            Else
             MsgBox "L'avvocato non vuole la fattura o non ha un numero di partita Iva valido.", vbInformation
        End If
    Else
        MsgBox "Il caricamento della testata non è andato a buon fine provare a rieseguire l'operazione!", vbCritical, "Attenzione"
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub IAnagraficForm_SelectCodiceAvvocato()
 TxtCodiceAvvocato.SetFocus
 'SendKeys "{Home}+{End}"
End Sub



