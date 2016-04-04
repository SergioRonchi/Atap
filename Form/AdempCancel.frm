VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form AdempCancel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Adempimenti di Cancelleria"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Tag             =   "Progressivo"
   Begin VB.Frame fraComandi 
      Height          =   900
      Left            =   20
      TabIndex        =   30
      Top             =   6000
      Width           =   9855
      Begin VB.CommandButton cmdPrint 
         Height          =   500
         Left            =   4080
         Picture         =   "AdempCancel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Stampa Schermata"
         Top             =   270
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "Ri&cerca Adempimenti"
         Height          =   500
         Left            =   2070
         TabIndex        =   28
         Top             =   270
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   500
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   1860
      End
      Begin VB.CommandButton CmdAnnulla 
         Caption         =   "E&sci"
         Height          =   500
         Left            =   7920
         TabIndex        =   29
         Top             =   270
         Width           =   1860
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   500
         Left            =   6000
         TabIndex        =   26
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame FraTop 
      Height          =   645
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2835
         TabIndex        =   2
         Top             =   225
         Width           =   285
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "XXX"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblCodiceA 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LblDescrCodAvv 
         Caption         =   "Descrizione:"
         DataField       =   "NOME"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Tag             =   "XXX"
         Top             =   240
         Width           =   6420
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   250
         Width           =   1200
      End
   End
   Begin VB.Frame fraMain 
      Height          =   5280
      Left            =   20
      TabIndex        =   33
      Top             =   720
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox txtSigla 
         DataField       =   "SIGLA"
         Height          =   285
         Left            =   3720
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
         Left            =   4920
         TabIndex        =   48
         Top             =   120
         Width           =   855
         Begin VB.Label LblProgressivo 
            Caption         =   "Progressivo : "
            Height          =   255
            Left            =   840
            TabIndex        =   50
            Top             =   -75
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label LblNumeroAtto 
            DataField       =   "Progressivo"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1800
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   930
         End
      End
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   680
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   4440
         Width           =   9615
         Begin VB.TextBox txtSiglaCH 
            DataField       =   "SIGLACH"
            Height          =   285
            Left            =   8760
            MaxLength       =   3
            TabIndex        =   24
            Tag             =   "Sigla Chiusura"
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox ChkAnnullo 
            Caption         =   "Check1"
            DataField       =   "Annullo"
            Height          =   240
            Left            =   1710
            TabIndex        =   25
            Tag             =   "PULISCI"
            ToolTipText     =   "Annulla"
            Top             =   360
            Width           =   240
         End
         Begin VB.CheckBox chkEvadi 
            Caption         =   "Check1"
            DataField       =   "CheckVisual"
            Height          =   240
            Left            =   3015
            TabIndex        =   23
            Tag             =   "PULISCI"
            ToolTipText     =   "Evadi"
            Top             =   0
            Width           =   240
         End
         Begin TDBDate6Ctl.TDBDate txtDataEvaso 
            DataField       =   "DataEvasionePratica"
            Height          =   255
            Left            =   1215
            TabIndex        =   22
            Top             =   0
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "AdempCancel.frx":014A
            Caption         =   "AdempCancel.frx":0262
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "AdempCancel.frx":02CE
            Keys            =   "AdempCancel.frx":02EC
            Spin            =   "AdempCancel.frx":034A
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
         Begin VB.Label LblAvvAdempAnn 
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
            Left            =   3615
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label Label5 
            Caption         =   "Sigla : "
            Height          =   255
            Left            =   8160
            TabIndex        =   59
            Top             =   360
            Width           =   510
         End
         Begin VB.Label LblDataEvaso 
            Caption         =   "Data Evasione : "
            Height          =   255
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label LblDescrEvaso 
            Caption         =   "Cancelleria evasa in data :"
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
            Left            =   3600
            TabIndex        =   54
            Top             =   0
            Visible         =   0   'False
            Width           =   5865
         End
         Begin VB.Label LblAnnullo 
            Caption         =   "Annulla adempimento: "
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   360
            Width           =   1590
         End
      End
      Begin TrueOleDBList80.TDBCombo cmbTribunale 
         DataField       =   "CodTribunaleApp"
         Height          =   315
         Left            =   6720
         TabIndex        =   5
         Tag             =   "necessario Tribunale"
         Top             =   140
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
         _PropDict       =   $"AdempCancel.frx":0372
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
      Begin TDBNumber6Ctl.TDBNumber txtSpese 
         DataField       =   "ImpSpese1E"
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   2070
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":03F9
         Caption         =   "AdempCancel.frx":0419
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":0485
         Keys            =   "AdempCancel.frx":04A3
         Spin            =   "AdempCancel.frx":04ED
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
      Begin VB.TextBox TxtAttivit‡Ric 
         DataField       =   "AttivitaRichiesta"
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   6
         Tag             =   "necessario Attivit‡ richesta"
         Top             =   480
         Width           =   7905
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese1"
         Height          =   285
         Index           =   1
         Left            =   675
         MaxLength       =   35
         TabIndex        =   8
         Top             =   2070
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese2"
         Height          =   285
         Index           =   2
         Left            =   675
         MaxLength       =   35
         TabIndex        =   10
         Top             =   2385
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese3"
         Height          =   285
         Index           =   3
         Left            =   675
         MaxLength       =   35
         TabIndex        =   12
         Top             =   2700
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese4"
         Height          =   285
         Index           =   4
         Left            =   675
         MaxLength       =   35
         TabIndex        =   14
         Top             =   3015
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese5"
         Height          =   285
         Index           =   5
         Left            =   675
         MaxLength       =   35
         TabIndex        =   16
         Top             =   3330
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpese 
         DataField       =   "DesrSpese6"
         Height          =   285
         Index           =   6
         Left            =   675
         MaxLength       =   35
         TabIndex        =   18
         Top             =   3645
         Width           =   2925
      End
      Begin VB.TextBox TxtMemo 
         DataField       =   "Memo"
         Height          =   1005
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   840
         Width           =   9600
      End
      Begin TDBNumber6Ctl.TDBNumber txtSpese 
         DataField       =   "ImpSpese2E"
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   11
         Top             =   2385
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0515
         Caption         =   "AdempCancel.frx":0535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":05A1
         Keys            =   "AdempCancel.frx":05BF
         Spin            =   "AdempCancel.frx":0609
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
         DataField       =   "ImpSpese3E"
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   13
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0631
         Caption         =   "AdempCancel.frx":0651
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":06BD
         Keys            =   "AdempCancel.frx":06DB
         Spin            =   "AdempCancel.frx":0725
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
         DataField       =   "ImpSpese4E"
         Height          =   285
         Index           =   4
         Left            =   3720
         TabIndex        =   15
         Top             =   3015
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":074D
         Caption         =   "AdempCancel.frx":076D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":07D9
         Keys            =   "AdempCancel.frx":07F7
         Spin            =   "AdempCancel.frx":0841
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
         DataField       =   "ImpSpese5E"
         Height          =   285
         Index           =   5
         Left            =   3720
         TabIndex        =   17
         Top             =   3330
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0869
         Caption         =   "AdempCancel.frx":0889
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":08F5
         Keys            =   "AdempCancel.frx":0913
         Spin            =   "AdempCancel.frx":095D
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
         DataField       =   "ImpSpese6E"
         Height          =   285
         Index           =   6
         Left            =   3720
         TabIndex        =   19
         Top             =   3645
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0985
         Caption         =   "AdempCancel.frx":09A5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":0A11
         Keys            =   "AdempCancel.frx":0A2F
         Spin            =   "AdempCancel.frx":0A79
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
      Begin TDBNumber6Ctl.TDBNumber txtDeposito 
         DataField       =   "ImpDepositoE"
         Height          =   285
         Left            =   7560
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0AA1
         Caption         =   "AdempCancel.frx":0AC1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":0B2D
         Keys            =   "AdempCancel.frx":0B4B
         Spin            =   "AdempCancel.frx":0B95
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
         Left            =   7560
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "AdempCancel.frx":0BBD
         Caption         =   "AdempCancel.frx":0BDD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":0C49
         Keys            =   "AdempCancel.frx":0C67
         Spin            =   "AdempCancel.frx":0CB1
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
      Begin TDBDate6Ctl.TDBDate txtDataReg 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Tag             =   "necessario Data Registrazione"
         Top             =   120
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calendar        =   "AdempCancel.frx":0CD9
         Caption         =   "AdempCancel.frx":0DF1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AdempCancel.frx":0E5D
         Keys            =   "AdempCancel.frx":0E7B
         Spin            =   "AdempCancel.frx":0ED9
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
      Begin VB.Label Label3 
         Caption         =   "Sigla : "
         Height          =   255
         Left            =   3120
         TabIndex        =   60
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "-"
         Height          =   255
         Left            =   7440
         TabIndex        =   58
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "-"
         Height          =   255
         Left            =   7440
         TabIndex        =   57
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   255
         Left            =   7440
         TabIndex        =   56
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label LblSpese 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   3720
         TabIndex        =   46
         Tag             =   "PULISCI"
         Top             =   4095
         Width           =   1185
      End
      Begin VB.Label LblSpeseVarie 
         Caption         =   "Spese Varie : "
         Height          =   255
         Left            =   3870
         TabIndex        =   45
         Top             =   1840
         Width           =   1005
      End
      Begin VB.Line Line1 
         X1              =   3600
         X2              =   5040
         Y1              =   4005
         Y2              =   4005
      End
      Begin VB.Label LblSommaSpese 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   7560
         TabIndex        =   44
         Top             =   2760
         Width           =   1185
      End
      Begin VB.Line Line2 
         X1              =   5940
         X2              =   8955
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label LblSpeseCalcSaldo 
         Caption         =   "Spese : "
         Height          =   255
         Left            =   6120
         TabIndex        =   43
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label LblTribunale 
         Caption         =   "Tribunale :"
         Height          =   250
         Left            =   5895
         TabIndex        =   42
         Top             =   140
         Width           =   825
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
         Height          =   250
         Left            =   7560
         TabIndex        =   41
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Label LblDescrSpeseVarie 
         Caption         =   "Descrizione : "
         Height          =   255
         Left            =   675
         TabIndex        =   39
         Top             =   1840
         Width           =   1050
      End
      Begin VB.Label LblSaldo 
         Caption         =   "Saldo : "
         Height          =   255
         Left            =   6120
         TabIndex        =   38
         Top             =   3375
         Width           =   510
      End
      Begin VB.Label LblCompetenze 
         Caption         =   "Competenze : "
         Height          =   255
         Left            =   6120
         TabIndex        =   37
         Top             =   2430
         Width           =   1140
      End
      Begin VB.Label LblDeposito 
         Caption         =   "Deposito : "
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   2080
         Width           =   870
      End
      Begin VB.Label LblAttivit‡Ric 
         Caption         =   "Attivit‡ Richiesta : "
         Height          =   255
         Left            =   225
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label LblDataReg 
         Caption         =   "Data Registrazione : "
         Height          =   250
         Left            =   225
         TabIndex        =   34
         Top             =   140
         Width           =   1455
      End
   End
End
Attribute VB_Name = "AdempCancel"
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
Implements IAnagraficForm
Implements IForm



Private Sub Timer1_Timer()
'CmdSalva.Enabled = Not IsRecordLocked("IDCod=" & m_ID, "ADEMPI")
End Sub

Private Sub ChkAnnullo_Click()
    If ChkAnnullo.value = Checked Then
        LblAvvAdempAnn.Visible = True
    Else
        LblAvvAdempAnn.Visible = False
    End If
End Sub

Private Sub chkEvadi_Click()
If chkEvadi = 1 Then
 If Not PassaLoad Then txtDataEvaso = Format(Date, "dd/mm/yyyy")
 LblDescrEvaso.Caption = "<< Cancelleria evasa"
 LblDescrEvaso.Visible = True
 txtSiglaCH.Tag = "necessario Sigla Chiusura"
Else
  txtDataEvaso = ""
  LblDescrEvaso.Visible = False
  txtSiglaCH.Tag = "Sigla Chiusura"
End If
End Sub

Private Sub cmbTribunale_Click()
 'Call InserisciPredefiniti
End Sub

Private Sub cmbTribunale_SelChange(Cancel As Integer)
 Call InserisciPredefiniti
End Sub

Private Sub CmdAnnulla_Click()
If CmdSalva.Enabled Then DeLockRecord m_ID, "ADEMPI"
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If
End Sub



Private Sub cmdPrint_Click()
  PrintForm
End Sub

Private Sub CmdRicerca_Click()
    Set moFrmRicerca = New FrmRicerca

    
    Set moFrmRicerca.frmCaller = Me
    moFrmRicerca.Titolo = "Ricerca Adempimenti"
    moFrmRicerca.tipo = "Ricerca"
    moFrmRicerca.Filtro = ""
    moFrmRicerca.DefaultOrder = "Order By DataRegistrazione DESC,ADEMPI.NumOrdinamento"
    moFrmRicerca.NCol = 6
    moFrmRicerca.PosizioneCodice = 7
    moFrmRicerca.Tabella = "ADEMPI"
    moFrmRicerca.Query = "SELECT CheckVisual AS Ev, CODAVV AS [Codice], " & _
                "Format(Mid(DataRegistrazione,7,2) & '/' & Mid(DataRegistrazione,5,2) & '/' & Mid(DataRegistrazione,1,4),'dd/mm/yyyy') As [Data Registrazione], " & _
                "AttivitaRichiesta as [Attivit‡],SIGLA as [Sigla Inserimento],SIGLACH as [Sigla chiusura],IDCod,Progressivo,CodTribunaleApp,DataEvasionePratica,Annullo,ADEMPI.NumOrdinamento FROM ADEMPI"
   
    Load moFrmRicerca
    

End Sub

Private Sub CmdRicercaA_Click()
On Error GoTo ErrHandler
   
   RicercaPerCodice Me, Azione
   txtDataReg = Date
 
    cmbTribunale = ""
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
        If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If

    Load FrmRicerca

End Sub

Private Sub CmdSalva_Click()
Dim Msg_Errore As String, contenuto1, contenuto2, contenuto3, contenuto4, contenuto5, Orario
Dim saved As Boolean
On Error GoTo ErroreSalvataggio
  
  If IsTableLocked("ADEMPI") Then
       MsgBox "La tabelle adempimenti Ë bloccata da un altro utente. Riprovare...", vbInformation
  Else
        'LockTable ("ADEMPI")
        
            saved = SalvaTutto(Me, "ADEMPI", sWhere)
            SaveSetting "ATAP", "Config", "Sigla", txtSigla.Text
            If Not moFrmRicerca Is Nothing Then
                moFrmRicerca.AggiornaGriglia
            End If
            If saved Then DeLockRecord m_ID, "ADEMPI"
            'DelockTable ("ADEMPI")
            TxtCodiceAvvocato.SetFocus
            
        End If
  
Exit Sub

ErroreSalvataggio:
    
    
    If CmdSalva.Caption = "&Modifica" Then
        Msg_Errore = "Errore durante la modifica di un adempimento cancelleria "
    Else
        Msg_Errore = "Errore durante il salavataggio di un adempimento cancelleria "
    End If
    Msg_Errore = Msg_Errore & " - numero: " & err & " - riga: " & Erl & " - messaggio: " & Error(err)
    Orario = (Date & " " & Time)
    contenuto1 = LblCodiceA.Caption & " " & TxtAttivit‡Ric.Text & " " & cmbTribunale.Text
    contenuto2 = "Data Evasione: " & txtDataEvaso.Text & " Data Registrazione: " & txtDataReg.Text & " " & TxtAttivit‡Ric.Text
    
    contenuto3 = TxtDescrSpese(1).Text & " " & txtSpese(1).Text & " " & TxtDescrSpese(2).Text & " " & txtSpese(2).Text & " " & TxtDescrSpese(3).Text & " " & txtSpese(3).Text
    contenuto4 = TxtDescrSpese(4).Text & " " & txtSpese(4).Text & " " & TxtDescrSpese(5).Text & " " & txtSpese(5).Text & " " & TxtDescrSpese(6).Text & " " & txtSpese(6).Text
    contenuto5 = txtCompetenze.Text & " " & txtDeposito.Text
    
    ErrLogFile "ErroriAtap.txt", Msg_Errore, contenuto1, contenuto2, contenuto3, contenuto4, contenuto5
    
 
    

End Sub
Private Function okdata() As Boolean
 okdata = TxtAttivit‡Ric.Text <> ""
End Function

Private Sub Form_Load()
    Me.Move 0, 0
    Azione = TipoAzione.Vuoto
    Call TipoMaschera(Me, Azione)
    PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", , , "DescrizioneTribunale"
    txtDataReg.MaxDate = Now + 30
    
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If
End Sub


Private Property Get IForm_IsLoading() As Boolean
IForm_IsLoading = PassaLoad
End Property

Private Property Let IForm_IsLoading(RHS As Boolean)
 PassaLoad = RHS
End Property

Private Sub IForm_RisRicerca()
'apri adempimento
Dim SQL As String
Dim rs As ADODB.Recordset

Set rs = newAdoRs
PassaLoad = True
TxtCodiceAvvocato.SetFocus
SQL = "SELECT ADEMPI.CODAVV, " & _
      "( Mid(ADEMPI.DataRegistrazione,7,2) & '/' & Mid(ADEMPI.DataRegistrazione,5,2)& '/' & Mid(ADEMPI.DataRegistrazione,1,4)) As DataRegistrazione, " & _
      "ADEMPI.Progressivo, ADEMPI.CodTribunaleApp, AnagraficaAvvocati.NOME, AnagraficaAvvocati.NumOrdinamento, " & _
      "ADEMPI.ImpDepositoE, ADEMPI.ImpSpese1E, ADEMPI.DesrSpese1, " & _
      "ADEMPI.ImpSpese2E, ADEMPI.DesrSpese2, ADEMPI.ImpSpese3E, ADEMPI.DesrSpese3, ADEMPI.ImpSpese4E, " & _
      "ADEMPI.DesrSpese4, ADEMPI.ImpSpese5E, ADEMPI.DesrSpese5, ADEMPI.ImpSpese6E, ADEMPI.DesrSpese6, " & _
      "ADEMPI.ImpCompetenzeE, ADEMPI.ImpSaldoE, " & _
      "( Mid(ADEMPI.DataEvasionePratica,7,2) & '/' & Mid(ADEMPI.DataEvasionePratica,5,2)& '/' & Mid(ADEMPI.DataEvasionePratica,1,4)) As DataEvasionePratica, ADEMPI.Annullo,ADEMPI.CheckVisual, " & _
      "ADEMPI.AttivitaRichiesta, ADEMPI.Memo,SIGLA,SIGLACH, ADEMPI.IDCod " & _
      "FROM (ADEMPI INNER JOIN AnagraficaAvvocati ON ADEMPI.CODAVV = AnagraficaAvvocati.CODAVV) INNER JOIN TribunaliAppartenenza ON ADEMPI.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale " & _
      "WHERE " & sWhere

rs.Open SQL, g_Settings.DBConnection
m_ID = -1
If Not rs.EOF Then
  
   Caricacampi Me, rs
   Azione = TipoAzione.Modifica
   Call TipoMaschera(Me, Azione)
    m_ID = rs("IDCod")
   If IsRecordLocked("IDCod=" & m_ID, "ADEMPI") Then
      CmdSalva.Enabled = False
     Else
      CmdSalva.Enabled = True
      LockRecord m_ID, "ADEMPI"
   End If
 Else
    MsgBox "Il caricamento non Ë andato a buon fine:" & vbCrLf & "potrebbe non essere presente la Cassetta o il Tribunale corrispondente", vbCritical, "Attenzione"
End If




  PassaLoad = False

End Sub

Private Sub IForm_SetFocus()
 Me.SetFocus
End Sub

Private Property Let IForm_Where(RHS As String)
 sWhere = RHS
End Property

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

Public Sub CalcolaSaldo()

Dim saldo As Double
Dim spese As Double
Dim i As Integer

    saldo = 0
    spese = 0
    
    For i = 1 To txtSpese.Count
      spese = spese + txtSpese(i).value
    Next
    saldo = txtDeposito.value

    saldo = saldo - txtCompetenze.value

    
    saldo = saldo - spese
    
        saldo = Format(saldo, "##,##0.00")
        spese = Format(spese, "##,##0.00")
    LblSommaSpese.Caption = Format(spese, "##,##0.00")
    
    
    LblSpese.Caption = Format(spese, "##,##0.00")
    
    formattaSaldo LblValSaldo, saldo

End Sub









Public Sub InserisciPredefiniti()
 Dim SQL As String
 Dim rs As ADODB.Recordset
 
 codTribunale = cmbTribunale.Columns(1).value
 
 SQL = "SELECT TribunaliAppartenenza.CodiceTribunale, Anticipi.PrezDepositoEuro, Anticipi.PrezCompetenzeEuro " & _
     "FROM Anticipi INNER JOIN TribunaliAppartenenza ON Anticipi.CodiceTribunale = TribunaliAppartenenza.CodiceTribunale " & _
     "WHERE Anticipi.CodiceAttivita='A' AND Anticipi.CodiceAlternativo='A' And TribunaliAppartenenza.CodiceTribunale='" & codTribunale & "'"
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

Private Sub txtDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txtSpese_Change(Index As Integer)
 Call CalcolaSaldo
End Sub

Private Sub txtSpese_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocato.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(CsCodAvv As String, oAzione As TipoAzione)
Dim rs As ADODB.Recordset
   m_ID = -1
    Azione = TipoAzione.Nuovo
    'Nuovo adempimento
    PassaLoad = True
    txtSigla = GetSetting("ATAP", "Config", "Sigla", "")
    Set rs = GetADORecordset("AnagraficaAvvocati", "CodAvv,Nome,numOrdinamento", "CodAvv='" & CsCodAvv & "'", g_Settings.DBConnection)
    
    If Not rs.EOF Then
     Call RiempiTestata(Me, rs)
     Call TipoMaschera(Me, Azione)
       
        
    Else
        MsgBox "Il caricamento della testata non Ë andato a buon fine provare a rieseguire l'operazione!", vbCritical, "Attenzione"
    End If
    rs.Close
    Set rs = Nothing
    PassaLoad = False

End Sub

Private Sub IAnagraficForm_SelectCodiceAvvocato()
 TxtCodiceAvvocato.SetFocus
 SendKeys "{Home}+{End}"
End Sub

