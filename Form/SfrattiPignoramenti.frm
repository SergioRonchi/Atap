VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form SfrattiPignoramenti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione sfratti e pignoramenti"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Tag             =   "NumeroAtto"
   Begin VB.PictureBox PictureUNEP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "SfrattiPignoramenti.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   61
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame fraComandi 
      Height          =   660
      Left            =   20
      TabIndex        =   40
      Top             =   5880
      Width           =   9855
      Begin VB.CommandButton cmdPrint 
         Height          =   450
         Left            =   3960
         Picture         =   "SfrattiPignoramenti.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Stampa Schermata"
         Top             =   150
         Width           =   1860
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   450
         Left            =   5880
         TabIndex        =   25
         Top             =   150
         Width           =   1860
      End
      Begin VB.CommandButton CmdAnnulla 
         Caption         =   "Esci"
         Height          =   450
         Left            =   7800
         TabIndex        =   26
         Top             =   150
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "Ri&cerca Sfratti "
         Height          =   450
         Left            =   2040
         TabIndex        =   23
         Top             =   150
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   450
         Left            =   120
         TabIndex        =   22
         Top             =   150
         Width           =   1860
      End
   End
   Begin VB.Frame fraMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   30
      Top             =   720
      Width           =   9855
      Begin VB.TextBox txtCronologico 
         DataField       =   "Crono"
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
         Height          =   285
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Cronologico"
         Top             =   480
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   3720
         Top             =   2160
      End
      Begin VB.TextBox txtSigla 
         DataField       =   "SIGLA"
         Height          =   285
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "necessario Sigla Inserimento"
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   47
         Top             =   120
         Width           =   5175
         Begin VB.Label LblAtto 
            Caption         =   "Numero atto : "
            Height          =   255
            Left            =   -600
            TabIndex        =   50
            Top             =   0
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LblNumeroAtto 
            DataField       =   "NumeroAtto"
            Height          =   255
            Left            =   1080
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin VB.TextBox TxtLocalita1 
         DataField       =   "Localita1"
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1440
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpeseVarie 
         DataField       =   "DesrSpese"
         Height          =   285
         Left            =   6480
         MaxLength       =   35
         TabIndex        =   17
         Top             =   2925
         Width           =   2925
      End
      Begin VB.TextBox TxtParte2 
         DataField       =   "Parte2"
         Height          =   285
         Left            =   6600
         MaxLength       =   35
         TabIndex        =   9
         Tag             =   "necessario Parte 2"
         Top             =   1155
         Width           =   3045
      End
      Begin VB.TextBox TxtParte1 
         DataField       =   "Parte1"
         Height          =   285
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   8
         Tag             =   "necessario Parte 1"
         Top             =   1155
         Width           =   2925
      End
      Begin TDBDate6Ctl.TDBDate txtDataReg 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Tag             =   "necessario Data Registrazione"
         Top             =   120
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calendar        =   "SfrattiPignoramenti.frx":058C
         Caption         =   "SfrattiPignoramenti.frx":06A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0710
         Keys            =   "SfrattiPignoramenti.frx":072E
         Spin            =   "SfrattiPignoramenti.frx":078C
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
      Begin TDBNumber6Ctl.TDBNumber txtDeposito 
         DataField       =   "ImpDepositoE"
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "SfrattiPignoramenti.frx":07B4
         Caption         =   "SfrattiPignoramenti.frx":07D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0840
         Keys            =   "SfrattiPignoramenti.frx":085E
         Spin            =   "SfrattiPignoramenti.frx":08A8
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
      Begin TDBNumber6Ctl.TDBNumber txtSpese 
         DataField       =   "ImpSpeseE"
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "SfrattiPignoramenti.frx":08D0
         Caption         =   "SfrattiPignoramenti.frx":08F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":095C
         Keys            =   "SfrattiPignoramenti.frx":097A
         Spin            =   "SfrattiPignoramenti.frx":09C4
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
      Begin TDBNumber6Ctl.TDBNumber txtCompetenze 
         DataField       =   "ImpCompetenzeE"
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "SfrattiPignoramenti.frx":09EC
         Caption         =   "SfrattiPignoramenti.frx":0A0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0A78
         Keys            =   "SfrattiPignoramenti.frx":0A96
         Spin            =   "SfrattiPignoramenti.frx":0AE0
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
      Begin TDBNumber6Ctl.TDBNumber txtVarie 
         DataField       =   "ImpVarieE"
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   2880
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "SfrattiPignoramenti.frx":0B08
         Caption         =   "SfrattiPignoramenti.frx":0B28
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0B94
         Keys            =   "SfrattiPignoramenti.frx":0BB2
         Spin            =   "SfrattiPignoramenti.frx":0BFC
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
      Begin TDBDate6Ctl.TDBDate TxtDataPresentazione 
         DataField       =   "DataPresentazione"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "SfrattiPignoramenti.frx":0C24
         Caption         =   "SfrattiPignoramenti.frx":0D3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0DA8
         Keys            =   "SfrattiPignoramenti.frx":0DC6
         Spin            =   "SfrattiPignoramenti.frx":0E24
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
      Begin TDBDate6Ctl.TDBDate TxtDataRestituzione 
         DataField       =   "DataRestituzione"
         Height          =   255
         Left            =   6600
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "SfrattiPignoramenti.frx":0E4C
         Caption         =   "SfrattiPignoramenti.frx":0F64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "SfrattiPignoramenti.frx":0FD0
         Keys            =   "SfrattiPignoramenti.frx":0FEE
         Spin            =   "SfrattiPignoramenti.frx":104C
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
         Left            =   1680
         TabIndex        =   6
         Tag             =   "necessario Tribunale"
         Top             =   795
         Width           =   2895
         _ExtentX        =   5106
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
         AutoCompletion  =   -1  'True
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
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
         AutoDropdown    =   -1  'True
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"SfrattiPignoramenti.frx":1074
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
      Begin TrueOleDBList80.TDBCombo CmbPignoramenti 
         DataField       =   "CodicePignoramenti"
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         Top             =   795
         Width           =   3015
         _ExtentX        =   5318
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
         AutoCompletion  =   -1  'True
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
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
         AutoDropdown    =   -1  'True
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"SfrattiPignoramenti.frx":10FB
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
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   4080
         Width           =   9135
         Begin VB.TextBox txtSiglaCH 
            DataField       =   "SIGLACH"
            Height          =   285
            Left            =   8280
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Sigla Chiusura"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox ChkAnnullo 
            Caption         =   "Check1"
            DataField       =   "Annullo"
            Height          =   240
            Left            =   2520
            TabIndex        =   20
            Tag             =   "PULISCI"
            Top             =   600
            Width           =   240
         End
         Begin VB.CheckBox chkEvadi 
            Caption         =   "Check1"
            DataField       =   "CheckVisual"
            Height          =   240
            Left            =   3375
            TabIndex        =   19
            ToolTipText     =   "Evadi"
            Top             =   120
            Width           =   240
         End
         Begin TDBDate6Ctl.TDBDate txtDataEvaso 
            DataField       =   "DataEvasionePratica"
            Height          =   255
            Left            =   1335
            TabIndex        =   18
            Top             =   120
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "SfrattiPignoramenti.frx":1182
            Caption         =   "SfrattiPignoramenti.frx":129A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "SfrattiPignoramenti.frx":1306
            Keys            =   "SfrattiPignoramenti.frx":1324
            Spin            =   "SfrattiPignoramenti.frx":1382
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   8454143
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
         Begin VB.Label Label6 
            Caption         =   "Sigla : "
            Height          =   255
            Left            =   7680
            TabIndex        =   60
            Top             =   600
            Width           =   510
         End
         Begin VB.Label LblAvvSfrPigAnn 
            Caption         =   "ANNULLATO"
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
            Left            =   3195
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label LblAnnullo 
            Caption         =   "Annulla sfratto / pignoramento: "
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label LblDescrEvaso 
            Caption         =   "Sfratto / pignoramento evaso in data :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3855
            TabIndex        =   52
            Top             =   120
            Visible         =   0   'False
            Width           =   3705
         End
         Begin VB.Label LblDataEvaso 
            Caption         =   "Data Evasione : "
            Height          =   255
            Left            =   15
            TabIndex        =   51
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Label lblCrono 
         Caption         =   "Cronologico : "
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
         Left            =   240
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Sigla : "
         Height          =   255
         Left            =   3000
         TabIndex        =   59
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         Height          =   255
         Left            =   1560
         TabIndex        =   58
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "-"
         Height          =   255
         Left            =   1560
         TabIndex        =   57
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "-"
         Height          =   255
         Left            =   1560
         TabIndex        =   56
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   2160
         Width           =   135
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   3120
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label LblLocalita1 
         Caption         =   "Località : "
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1545
         Width           =   720
      End
      Begin VB.Label LblPignoramenti 
         Caption         =   "Pignoramenti :"
         Height          =   255
         Left            =   5400
         TabIndex        =   45
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label LblTribunale 
         Caption         =   "Tribunale :"
         Height          =   255
         Left            =   225
         TabIndex        =   44
         Top             =   855
         Width           =   825
      End
      Begin VB.Label LblDataRestituzione 
         Caption         =   "Restituzione : "
         Height          =   495
         Left            =   5400
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label LblDataPresentazione 
         Caption         =   "Presentazione: "
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label LblParte2 
         Caption         =   "Parte 2: "
         Height          =   255
         Left            =   5400
         TabIndex        =   41
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label LblDataReg 
         Caption         =   "Registrazione : "
         Height          =   375
         Left            =   225
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LblParte1 
         Caption         =   "Parte 1: "
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label LblDeposito 
         Caption         =   "Deposito : "
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label LblSpese 
         Caption         =   "Spese : "
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label LblCompetenze 
         Caption         =   "Competenze : "
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label LblSaldo 
         Caption         =   "Saldo : "
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3720
         Width           =   510
      End
      Begin VB.Label LblVarie 
         Caption         =   "Varie : "
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label LblDescrSpeseVarie 
         Caption         =   "Descrizione : "
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label LblValSaldo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "ImpSaldoE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   3720
         Width           =   1260
      End
   End
   Begin VB.Frame fraTop 
      Height          =   645
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.TextBox TxtCodiceAvvocato 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "XXX"
         Top             =   240
         Width           =   1290
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   285
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cassetta :"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label LblDescrCodAvv 
         Caption         =   "NOME"
         DataField       =   "NOME"
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Tag             =   "XXX"
         Top             =   240
         Width           =   5580
      End
      Begin VB.Label LblCodiceA 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   270
         Width           =   960
      End
   End
End
Attribute VB_Name = "SfrattiPignoramenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numOrdinamento As Integer
Dim codTribunale As String

Dim PassaLoad As Boolean
Public Azione As TipoAzione
Private sWhere As String
Private moFrmRicerca As FrmRicerca
Private m_ID As Long

Public isUnep As Boolean


Implements IAnagraficForm
Implements IForm

Private Sub chkEvadi_Click()
If chkEvadi = 1 Then
 If Not PassaLoad Then txtDataEvaso = Format(Date, "dd/mm/yyyy")
 LblDescrEvaso.Caption = "<< Sfratto/Pignoramento evaso"
 LblDescrEvaso.Visible = True
 txtSiglaCH.Tag = "necessario Sigla Chiusura"
Else
  txtDataEvaso = ""
  LblDescrEvaso.Visible = False
  txtSiglaCH.Tag = "Sigla Chiusura"
End If

End Sub




Private Sub CmbPignoramenti_SelChange(Cancel As Integer)
If Not PassaLoad Then InserisciPredefiniti
End Sub

Private Sub cmbTribunale_SelChange(Cancel As Integer)
If Not PassaLoad Then InserisciPredefiniti
End Sub

Private Sub CmdAnnulla_Click()
If CmdSalva.Enabled Then DeLockRecord m_ID, getCurrentTable
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If
End Sub
Private Sub cmdPrint_Click()
  PrintForm
End Sub

Private Sub CmdRicercaA_Click()
On Error GoTo ErrHandler
    'Ricerca Avvocato
   
   RicercaPerCodice Me, Azione
   txtDataReg = Date
   cmbTribunale = ""
   CmbPignoramenti = ""
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
    
    If isUnep Then
      FrmRicerca.Filtro = FrmRicerca.Filtro & " AND NOT (CODAVV LIKE '525%' OR CODAVV LIKE '393%')"
    End If
        If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If

    Load FrmRicerca

End Sub

Private Sub CmdRicerca_Click()
   Set moFrmRicerca = New FrmRicerca
    
    Set moFrmRicerca.frmCaller = Me
        moFrmRicerca.tipo = "Ricerca"
    moFrmRicerca.Filtro = ""

    moFrmRicerca.Titolo = "Ricerca Sfratti/Pignoramenti"
    moFrmRicerca.DefaultOrder = "Order By DataRegistrazione DESC, NumOrdinamento"
    moFrmRicerca.NCol = IIf(isUnep, 8, 7)
    moFrmRicerca.PosizioneCodice = 9
    moFrmRicerca.Tabella = getCurrentTable()
     moFrmRicerca.isUnep = isUnep
     
    moFrmRicerca.Query = "SELECT CheckVisual AS Ev, CODAVV AS [Codice], " & _
                "Format(Mid(DataRegistrazione,7,2) & '/' & Mid(DataRegistrazione,5,2) & '/' & Mid(DataRegistrazione,1,4),'dd/mm/yyyy') As [Data Registrazione], " & _
                "Parte1, Parte2,SIGLA as [Sigla Inserimento],SIGLACH as [Sigla chiusura],Crono, IDCod,NumeroAtto, CodTribunaleApp, DataEvasionePratica,Annullo,NumOrdinamento " & _
                "FROM " & getCurrentTable() & " "
   
    
    Load moFrmRicerca
End Sub
Private Function getCurrentTable() As String
If isUnep Then
 getCurrentTable = "SFRATTI_UNEP"
Else
 getCurrentTable = "SFRATTI"
End If
End Function
Private Sub CmdSalva_Click()
Dim Msg_Errore As String
Dim saved As Boolean
On Error GoTo ErroreSalvataggio



  If IsTableLocked(getCurrentTable()) Then
       MsgBox "La tabelle sfratti è bloccata da un altro utente. Riprovare...", vbInformation
  Else
        'LockTable ("SFRATTI")
        SaveSetting "ATAP", "Config", "Sigla", txtSigla.Text
        saved = SalvaTutto(Me, getCurrentTable(), sWhere, True)
        
        If Not moFrmRicerca Is Nothing Then
            moFrmRicerca.AggiornaGriglia
        End If

        If saved Then DeLockRecord m_ID, getCurrentTable()
        'DelockTable ("SFRATTI")
        'TxtCodiceAvvocato.SetFocus
  End If


Exit Sub

ErroreSalvataggio:

    If CmdSalva.Caption = "&Modifica" Then
        Msg_Errore = "Errore durante la modifica di uno sfratto/pignoramento "
    Else
        Msg_Errore = "Errore durante il salavataggio di uno sfratto/pignoramento "
    End If
    Msg_Errore = Msg_Errore & " - numero: " & err & " - riga: " & Erl & " - messaggio: " & Error(err)

    
    ErrLogFile "ErroriAtap.txt", Msg_Errore



End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    Azione = TipoAzione.Vuoto
    Call TipoMaschera(Me, Azione)
    txtDataReg.MaxDate = Now + 30
    PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", , , "DescrizioneTribunale"
    PopolaTDBCombo CmbPignoramenti, "Pignoramenti", "Descrizione", "Codice"
    PictureUNEP.Visible = False
    If isUnep Then
       Me.Caption = Me.Caption + " : UNEP"
       LblCompetenze.Caption = "Proporzionale"
       PictureUNEP.Visible = True
       lblCrono.Visible = True
       txtCronologico.Visible = True
       
       Dim r
       r = cmbTribunale.Columns(1).Find("UNEP", dblSeekEQ, True)
       If Not IsNull(r) Then cmbTribunale.Bookmark = r
       cmbTribunale.BoundText = cmbTribunale.Columns(0).value
       cmbTribunale.Enabled = False
    End If
End Sub




Private Property Let IForm_IsLoading(RHS As Boolean)
 PassaLoad = RHS
End Property
Private Property Get IForm_IsLoading() As Boolean
IForm_IsLoading = PassaLoad
End Property
Private Sub IForm_SetFocus()
 Me.SetFocus
End Sub
Private Sub IForm_RisRicerca()
 Dim SQL As String
Dim rs As ADODB.Recordset
 PassaLoad = True
 
    LblAtto.Visible = True
    LblNumeroAtto.Visible = True
    TxtCodiceAvvocato.Visible = False
    CmdRicercaA.Visible = False
    
Set rs = newAdoRs
PassaLoad = True
SQL = "SELECT " & getCurrentTable() & ".CODAVV, " & _
      "( Mid(DataRegistrazione,7,2) & '/' & Mid(DataRegistrazione,5,2)& '/' & Mid(DataRegistrazione,1,4)) As DataRegistrazione, " & _
      "( Mid(DataPresentazione,7,2) & '/' & Mid(DataPresentazione,5,2)& '/' & Mid(DataPresentazione,1,4)) As DataPresentazione, " & _
      "( Mid(DataRestituzione,7,2) & '/' & Mid(DataRestituzione,5,2)& '/' & Mid(DataRestituzione,1,4)) As DataRestituzione, " & _
      "NumeroAtto, CodicePignoramenti,CodTribunaleApp, AnagraficaAvvocati.NOME, AnagraficaAvvocati.NumOrdinamento, " & _
      "ImpDepositoE, ImpSpeseE, DesrSpese, " & _
      "ImpCompetenzeE, ImpSaldoE, " & _
      "ImpVarieE,Parte1,Localita1,Parte2, " & _
      "( Mid(DataEvasionePratica,7,2) & '/' & Mid(DataEvasionePratica,5,2)& '/' & Mid(DataEvasionePratica,1,4)) As DataEvasionePratica, Annullo,CheckVisual, " & _
      getCurrentTable() & ".NumOrdinamento,SIGLA,SIGLACH, IDCod, Crono " & _
      "FROM (" & getCurrentTable() & " INNER JOIN AnagraficaAvvocati ON " & getCurrentTable() & ".CODAVV = AnagraficaAvvocati.CODAVV) INNER JOIN TribunaliAppartenenza ON " & getCurrentTable() & ".CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale " & _
      "WHERE " & sWhere
rs.Open SQL, g_Settings.DBConnection
m_ID = -1
If Not rs.EOF Then
   
   Caricacampi Me, rs
   Azione = TipoAzione.Modifica
   Call TipoMaschera(Me, Azione)
         m_ID = rs("IDCod")
   If IsRecordLocked("IDCod=" & m_ID, getCurrentTable()) Then
      CmdSalva.Enabled = False
     Else
      CmdSalva.Enabled = True
      LockRecord m_ID, getCurrentTable()
   End If
 Else
    MsgBox "Il caricamento non è andato a buon fine:" & vbCrLf & "potrebbe non essere presente la Cassetta o il Tribunale corrispondente", vbCritical, "Attenzione"
End If

 
 PassaLoad = False

End Sub

Private Property Let IForm_Where(RHS As String)
 sWhere = RHS
End Property

Private Sub Timer1_Timer()
'CmdSalva.Enabled = Not IsRecordLocked("IDCod=" & m_ID, "SFRATTI")
End Sub

Private Sub TxtCodiceAvvocato_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CmdRicercaA_Click
End Sub

Private Sub TxtCompetenze_Change()

    Call CalcolaSaldo
End Sub


Private Sub txtCompetenze_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub
Private Sub TxtDeposito_Change()
    Call CalcolaSaldo
End Sub
Private Sub txtDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txtSpese_Change()
    Call CalcolaSaldo
End Sub

Private Sub txtSpese_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub TxtVarie_Change()
    Call CalcolaSaldo
End Sub

Public Sub CalcolaSaldo()

Dim saldo As Double

    saldo = 0
    saldo = txtDeposito.value
    saldo = saldo - txtSpese.value
    saldo = saldo - txtVarie.value
    saldo = saldo - txtCompetenze.value
    

    formattaSaldo LblValSaldo, saldo

End Sub



Private Sub ChkAnnullo_Click()
    If ChkAnnullo.value = Checked Then
        LblAvvSfrPigAnn.Visible = True
    Else
        LblAvvSfrPigAnn.Visible = False
    End If
End Sub

Public Sub InserisciPredefiniti()
 Dim SQL As String
 Dim rs As ADODB.Recordset
 Dim codPigno, codTribunale
 codTribunale = cmbTribunale.Columns(1).value
 codPigno = CmbPignoramenti.Columns(1).value
 SQL = "SELECT TribunaliAppartenenza.CodiceTribunale, Anticipi.PrezDepositoEuro, Anticipi.PrezCompetenzeEuro " & _
     "FROM Anticipi INNER JOIN TribunaliAppartenenza ON Anticipi.CodiceTribunale = TribunaliAppartenenza.CodiceTribunale " & _
     "WHERE Anticipi.CodiceAttivita='S' AND Anticipi.CodiceAlternativo='" & codPigno & "' And TribunaliAppartenenza.CodiceTribunale='" & codTribunale & "'"
  Set rs = newAdoRs
  
  rs.Open SQL, g_Settings.DBConnection
  If Not rs.EOF Then
           txtDeposito = rs!PrezDepositoEuro
           txtCompetenze = rs!PrezCompetenzeEuro
      Else
           txtDeposito.value = 0
           txtCompetenze.value = 0
  End If
  
  

  rs.Close

End Sub


Private Sub txtVarie_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
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
    
    
    If isUnep Then
       Dim r
       r = cmbTribunale.Columns(1).Find("UNEP", dblSeekEQ, True)
       If Not IsNull(r) Then cmbTribunale.Bookmark = r
       cmbTribunale.BoundText = cmbTribunale.Columns(0).value
       cmbTribunale.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub IAnagraficForm_SelectCodiceAvvocato()
 TxtCodiceAvvocato.SetFocus
 SendKeys "{Home}+{End}"
End Sub

