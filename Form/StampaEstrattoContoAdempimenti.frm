VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaEstrattoContoAdempimenti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Estratto Conto Adempimenti"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   3840
      TabIndex        =   11
      Top             =   3840
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2400
      TabIndex        =   9
      Top             =   3840
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   3585
      Left            =   75
      TabIndex        =   7
      Top             =   120
      Width           =   5205
      Begin VB.Frame FrmProvvisoria 
         Height          =   2115
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4920
         Begin VB.CommandButton CmdRicercaA 
            Caption         =   "->"
            Height          =   285
            Left            =   2760
            TabIndex        =   3
            Top             =   810
            Width           =   330
         End
         Begin VB.TextBox TxtCodiceAvvocato 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   2
            Top             =   810
            Width           =   1350
         End
         Begin VB.CommandButton CmdRicercaAnag 
            Caption         =   "&Ricerca Anagrafica"
            Height          =   525
            Left            =   3555
            TabIndex        =   4
            Top             =   810
            Width           =   1215
         End
         Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
            DataField       =   "DataRegistrazione"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoContoAdempimenti.frx":0000
            Caption         =   "StampaEstrattoContoAdempimenti.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoContoAdempimenti.frx":0184
            Keys            =   "StampaEstrattoContoAdempimenti.frx":01A2
            Spin            =   "StampaEstrattoContoAdempimenti.frx":0200
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
            MinDate         =   -657434
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
         Begin TDBDate6Ctl.TDBDate TxtRicDataFin 
            DataField       =   "DataRegistrazione"
            Height          =   255
            Left            =   2400
            TabIndex        =   1
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoContoAdempimenti.frx":0228
            Caption         =   "StampaEstrattoContoAdempimenti.frx":0340
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoContoAdempimenti.frx":03AC
            Keys            =   "StampaEstrattoContoAdempimenti.frx":03CA
            Spin            =   "StampaEstrattoContoAdempimenti.frx":0428
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
            MinDate         =   -657434
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
         Begin VB.Label LblRicDataIn 
            Caption         =   "Data Inizio :"
            Height          =   285
            Left            =   135
            TabIndex        =   17
            Top             =   120
            Width           =   870
         End
         Begin VB.Label LblRicDataFin 
            Caption         =   "Data Fine :"
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   120
            Width           =   825
         End
         Begin VB.Label LblDescr 
            Caption         =   "Descrizione:"
            Height          =   255
            Left            =   135
            TabIndex        =   15
            Top             =   1200
            Width           =   1110
         End
         Begin VB.Label LblCodAvvocato 
            Caption         =   "Cod. Cassetta:"
            Height          =   255
            Left            =   135
            TabIndex        =   14
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label LblDescrCodAvv 
            Caption         =   "TUTTE LE CASSETTE"
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   4545
         End
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Esporta Adempimenti"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2505
      End
      Begin VB.Frame FrmMetodoStampa 
         Caption         =   "Modalità Stampa"
         Height          =   645
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   4920
         Begin VB.OptionButton OptModSt 
            Caption         =   "Diretta"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   6
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton OptModSt 
            Caption         =   "Anteprima"
            Height          =   195
            Index           =   0
            Left            =   855
            TabIndex        =   5
            Top             =   270
            Value           =   -1  'True
            Width           =   1680
         End
      End
   End
   Begin Crystal.CrystalReport CRptEstratto 
      Left            =   4560
      Top             =   3855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaEstrattoContoAdempimenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TrasferimentoOK As Boolean
Private WithEvents moFilterManager As CFilterManager
Attribute moFilterManager.VB_VarHelpID = -1
Const prtTableName = "PrtEstrattoConto"

Private Sub CmdAnnulla_Click()
    Dim res As VbMsgBoxResult
    Dim OK As Boolean
    Dim nome As String
    nome = TxtCodiceAvvocato.Text
    If nome = "" Then nome = "COMPLETO"
    res = MsgBox("Vuoi cancellare gli Adempimenti stampati, " & vbCrLf & "trasferendoli in un archivio storico?", vbYesNo, "Attenzione")
    If (res = vbYes) Then
     OK = Trasferisci(g_Settings.StoricoECAdempiPath & "\EC_ADEMPI_" & Format(Now, "yyyymmddhhmm") & "_" & nome & ".mdb", Format(TxtRicDataIn.Text, "yyyymmdd"), Format(TxtRicDataFin.Text, "yyyymmdd"), False, Trim(TxtCodiceAvvocato.Text), "A")
     
    End If
End Sub


Private Sub CmdOK_Click()
    
    If IsPrtTableLocked("PrtEstrattoConto") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If

    LockPrtTable ("PrtEstrattoConto")
 
    
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, TxtCodiceAvvocato.Text, 1, 0, 0, 0, "N", False, 0
    If Not GetADORecordset("PrtEstrattoConto", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
      
          Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "EstrattoConto.rpt", 1, "Tipo='CANCELLERIE'")
    Else
        MsgBox "Nessun dato evaso! Impossibile creare l'Estratto Conto Adempimenti.", vbInformation, "Attenzione"
    End If
    DelockPrtTable ("PrtEstrattoConto")
End Sub

Private Sub Command1_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub
Private Sub moFilterManager_Validate(IsValid As Boolean)
   CmdOK.Enabled = IsValid
   CmdAnnulla.Enabled = IsValid
End Sub
Private Sub Form_Load()
   Set moFilterManager = New CFilterManager
   moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv
   
   Me.Move 400, 400
End Sub
