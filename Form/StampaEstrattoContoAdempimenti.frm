VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form StampaEstrattoContoAdempimenti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Estratto Conto Adempimenti"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   3840
      TabIndex        =   11
      Top             =   5280
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2400
      TabIndex        =   9
      Top             =   5280
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   5025
      Left            =   75
      TabIndex        =   7
      Top             =   120
      Width           =   5205
      Begin VB.Frame FrmProvvisoria 
         Height          =   3315
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
         Begin TrueOleDBList80.TDBCombo cmbTribunale 
            DataField       =   "CodTribunaleApp"
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Tag             =   "necessario Tribunale"
            Top             =   2520
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   0
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   315,213
            AutoSize        =   -1  'True
            GapHeight       =   30,047
            ListField       =   ""
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            AutoDropdown    =   0   'False
            RowTracking     =   -1  'True
            RightToLeft     =   0   'False
            RowMember       =   ""
            MouseIcon       =   0
            MouseIcon.vt    =   3
            MousePointer    =   0
            MatchEntryTimeout=   2000
            OLEDragMode     =   0
            OLEDropMode     =   0
            AnimateWindow   =   1
            AnimateWindowDirection=   0
            AnimateWindowTime=   200
            AnimateWindowClose=   0
            DropdownPosition=   0
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            AddItemSeparator=   ";"
            _PropDict       =   $"StampaEstrattoContoAdempimenti.frx":0450
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Caption         =   "Tribunale"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   975
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
         Top             =   4560
         Width           =   2505
      End
      Begin VB.Frame FrmMetodoStampa 
         Caption         =   "Modalità Stampa"
         Height          =   645
         Left            =   120
         TabIndex        =   10
         Top             =   3720
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
      Top             =   5295
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
    Dim avvocatoScelto As String

avvocatoScelto = Trim(TxtCodiceAvvocato.Text)
    Dim avvocatiEstratti As New AvvocatiPerEstratto

If avvocatoScelto = "" Then
  avvocatiEstratti.Tutti = True
 Else
 avvocatiEstratti.Lista.Add avvocatoScelto
End If
    
    If (res = vbYes) Then
     OK = Trasferisci(g_Settings.StoricoECAdempiPath & "\EC_ADEMPI_" & Format(Now, "yyyymmddhhmm") & "_" & nome & ".mdb", Format(TxtRicDataIn.Text, "yyyymmdd"), Format(TxtRicDataFin.Text, "yyyymmdd"), False, avvocatiEstratti, "A")
     
    End If
End Sub


Private Sub CmdOK_Click()
  Dim avvocatiEstratti As AvvocatiPerEstratto
  Dim codTribunale As String
  Set avvocatiEstratti = GetAvvocatoSingoloPerEstratto(TxtCodiceAvvocato.Text)
   If Not IsDate(TxtRicDataIn.Text) Or Not IsDate(TxtRicDataFin.Text) Then
    MsgBox "Inserire l'intervallo di date", vbOKOnly + vbCritical
    Exit Sub
  End If
    If IsPrtTableLocked("PrtEstrattoConto") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If

    LockPrtTable ("PrtEstrattoConto")
 
    If Not IsNull(cmbTribunale.SelectedItem) Then
        codTribunale = cmbTribunale.Columns(1).value
        If codTribunale = "XXALLXX" Then codTribunale = ""
    End If
    
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, 1, 0, 0, 0, "N", False, 0, codTribunale
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
   cmdAnnulla.Enabled = IsValid
End Sub
Private Sub Form_Load()
   Set moFilterManager = New CFilterManager
   PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", True
   


   moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv
   
   Me.Move 400, 400
End Sub
