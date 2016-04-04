VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form Anticipi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Anticipi"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmAnticipi 
      Height          =   675
      Left            =   50
      TabIndex        =   11
      Top             =   0
      Width           =   10770
      Begin TrueOleDBList80.TDBCombo CmbAttivita 
         Bindings        =   "Anticipi.frx":0000
         DataField       =   "CodiceAttivita"
         Height          =   315
         Left            =   960
         TabIndex        =   32
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
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
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         AddItemSeparator=   ";"
         _PropDict       =   $"Anticipi.frx":0019
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "Ricerca"
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin TrueOleDBList80.TDBCombo cmbTribunale 
         DataField       =   "CodiceTribunale"
         Height          =   315
         Left            =   5520
         TabIndex        =   33
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
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
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         AddItemSeparator=   ";"
         _PropDict       =   $"Anticipi.frx":00A0
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.Label LblTribunale 
         Caption         =   "Tribunale :"
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblAttivita 
         Caption         =   "Attività :"
         Height          =   285
         Left            =   270
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame FrmPrezzi 
      Height          =   1920
      Left            =   50
      TabIndex        =   14
      Top             =   650
      Width           =   10755
      Begin VB.TextBox TxtCodice 
         DataField       =   "CodiceAnticipi"
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   2
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox TxtDescrizione 
         DataField       =   "Descrizione"
         Height          =   285
         Left            =   7560
         MaxLength       =   35
         TabIndex        =   3
         Top             =   240
         Width           =   3000
      End
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezDepositoEuro"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":0127
         Caption         =   "Anticipi.frx":0147
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":01B3
         Keys            =   "Anticipi.frx":01D1
         Spin            =   "Anticipi.frx":021B
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
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezFotocopieEuro"
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":0243
         Caption         =   "Anticipi.frx":0263
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":02CF
         Keys            =   "Anticipi.frx":02ED
         Spin            =   "Anticipi.frx":0337
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
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezCompetenzeEuro"
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":035F
         Caption         =   "Anticipi.frx":037F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":03EB
         Keys            =   "Anticipi.frx":0409
         Spin            =   "Anticipi.frx":0453
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
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezFormulaEuro"
         Height          =   285
         Index           =   3
         Left            =   9360
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":047B
         Caption         =   "Anticipi.frx":049B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":0507
         Keys            =   "Anticipi.frx":0525
         Spin            =   "Anticipi.frx":056F
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
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezCancelleriaEuro"
         Height          =   285
         Index           =   5
         Left            =   5520
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":0597
         Caption         =   "Anticipi.frx":05B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":0623
         Keys            =   "Anticipi.frx":0641
         Spin            =   "Anticipi.frx":068B
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
      Begin TDBNumber6Ctl.TDBNumber txtPrezzo 
         DataField       =   "PrezMarcheEuro"
         Height          =   285
         Index           =   6
         Left            =   9360
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Anticipi.frx":06B3
         Caption         =   "Anticipi.frx":06D3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Anticipi.frx":073F
         Keys            =   "Anticipi.frx":075D
         Spin            =   "Anticipi.frx":07A7
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
      Begin TrueOleDBList80.TDBCombo CmbCodiceAlternativo 
         DataField       =   "CodiceAlternativo"
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Tag             =   "A"
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
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
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         AddItemSeparator=   ";"
         _PropDict       =   $"Anticipi.frx":07CF
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.Label LblCodice 
         Caption         =   "Codice :"
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label LblDescrizione 
         Caption         =   "Descrizione :"
         Height          =   285
         Left            =   6360
         TabIndex        =   27
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LblCodAlternativo 
         Caption         =   "Codice Alt"
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label LblAutorita 
         Height          =   285
         Left            =   4680
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Marche :"
         DataField       =   "Marche"
         Height          =   285
         Index           =   6
         Left            =   8520
         TabIndex        =   20
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Diritti Cancelleria :"
         DataField       =   "DirittiCancelleria"
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   19
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Fotocopie :"
         DataField       =   "Fotocopie"
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Formula :"
         DataField       =   "Formula"
         Height          =   285
         Index           =   3
         Left            =   8520
         TabIndex        =   17
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Competenze :"
         DataField       =   "Competenze"
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   16
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label LblPrezzo 
         Caption         =   "Deposito :"
         DataField       =   "Deposito"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   960
      End
   End
   Begin VB.Frame FrmRicercaAnticipi 
      Height          =   4635
      Left            =   50
      TabIndex        =   10
      Top             =   2580
      Width           =   10755
      Begin VSFlex8Ctl.VSFlexGrid flex 
         Height          =   4455
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   10695
         _cx             =   18865
         _cy             =   7858
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame FrmButtonZone 
      Height          =   825
      Left            =   50
      TabIndex        =   0
      Top             =   7320
      Width           =   10755
      Begin VB.CommandButton Command1 
         Caption         =   "E&sci"
         Height          =   500
         Left            =   9360
         TabIndex        =   31
         Top             =   200
         Width           =   1200
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "Annulla"
         Height          =   500
         Left            =   3000
         TabIndex        =   24
         Top             =   200
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton CmdAggiungi 
         Caption         =   "Nuovo"
         Height          =   500
         Left            =   360
         TabIndex        =   21
         Top             =   200
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   500
         Left            =   7860
         Picture         =   "Anticipi.frx":0856
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Stampa Schermata"
         Top             =   200
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   500
         Left            =   4320
         TabIndex        =   26
         Top             =   200
         Width           =   1200
      End
      Begin VB.CommandButton CmdElimina 
         Caption         =   "&Elimina"
         Height          =   500
         Left            =   1680
         TabIndex        =   23
         Top             =   200
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Anticipi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Query As String
Dim colTribunali As Collection
Dim colAttivita As Collection
Dim colPigno As Collection
Dim colAuto As Collection
Dim colAtto As Collection
Dim Azione As TipoAzione







Private Sub CmbCodiceAlternativo_Change()
LblAutorita.Caption = CmbCodiceAlternativo.Columns(0).value

End Sub

Private Sub CmbCodiceAlternativo_Click()
    CmbCodiceAlternativo_Change
End Sub

Private Sub CmdAggiungi_Click()
    Azione = TipoAzione.Nuovo
    pulisciCodice
    If CmbAttivita.Columns(1).value = "XXALLXX" Then
            MsgBox "Prima di aggiungere un anticipo si deve sceglie l'attivita!", vbInformation, "Attenzione"
            CmbAttivita.SetFocus
            Exit Sub
    End If
    If cmbTribunale.Columns(1).value = "XXALLXX" Then
            MsgBox "Prima di aggiungere un anticipo si deve sceglie un tribunale!", vbInformation, "Attenzione"
            cmbTribunale.SetFocus
            Exit Sub
    End If
        GestioneCodiceAlternativo
        VisualizzaPrezzi (CmbAttivita.Columns(1).value)
        FrmPrezzi.Visible = True
        BloccaSblocca (False)
End Sub
Private Sub BloccaSblocca(Sblocca As Boolean)
CmdAggiungi.Enabled = Sblocca
CmdElimina.Enabled = Sblocca

CmdAnnulla.Visible = Not Sblocca
FrmRicercaAnticipi.Enabled = Sblocca
FrmAnticipi.Enabled = Sblocca
End Sub

Private Sub CmdAnnulla_Click()
BloccaSblocca True
FrmPrezzi.Visible = False
End Sub

Private Sub CmdElimina_Click()
Dim app As String
Dim code As String
Dim r As Long
Dim Response As String
r = flex.row
code = flex.TextMatrix(r, 1)
app = flex.TextMatrix(r, 2)
Response = MsgBox("Vuoi eliminare il record  " & code & " - " & app & " ?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
If Response = vbYes Then    ' User chose Yes.
  g_Settings.DBConnection.Execute "DELETE FROM ANTICIPI WHERE CodiceAnticipi=" & code & ";"

  pulisci
  Visualizza False
  

 End If
 
End Sub

Private Sub CmdModifica_Click()
Edit
End Sub



Private Sub cmdPrint_Click()
  PrintForm
End Sub
Private Sub Ricerca()
Dim qry1 As String, qry2 As String
    If CmbAttivita.Columns(1).value <> "XXALLXX" Then qry1 = " AND (CodiceAttivita = '" & CmbAttivita.Columns(1).value & "')"
    If cmbTribunale.Columns(1).value <> "XXALLXX" Then qry2 = " AND (CodiceTribunale = '" & cmbTribunale.Columns(1).value & "')"
    Query = "SELECT CodiceAnticipi as codice, Descrizione, PrezDepositoEuro as Deposito, PrezCompetenzeEuro as Competenze, PrezFormulaEuro as formula, PrezFotocopieEuro as Fotocopie, PrezCancelleriaEuro as Cancelleria, PrezMarcheEuro as Marche FROM Anticipi "
    Query = Query & "WHERE 1=1 " & qry1 & qry2 & " ORDER BY CodiceAnticipi"

Visualizza False

End Sub

Private Sub CmdRicerca_Click()
Ricerca
End Sub

Private Sub CmdSalva_Click()
Dim sWhere As String
Dim Salvato As Boolean
Dim Response As String
If FrmPrezzi.Visible = True Then
    If TxtCodice.Text = "" Then
        MsgBox "Codice obbligatorio!", vbInformation, "Attenzione"
        TxtCodice.SetFocus
        Exit Sub
    End If
    If TxtDescrizione.Text = "" Then
        MsgBox "Descrizione obbligatoria!", vbInformation, "Attenzione"
        TxtDescrizione.SetFocus
        Exit Sub
    End If
    If CmbAttivita.Columns(1).value <> "A" Then
      If CmbCodiceAlternativo.Text = "" Then
         MsgBox LblCodAlternativo.Caption & " obbligatorio!", vbInformation, "Attenzione"
         CmbCodiceAlternativo.SetFocus
         Exit Sub
      End If
     End If
    If Azione = TipoAzione.Modifica Then
           Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
     Else
           Response = MsgBox("Vuoi salvare i dati inseriti?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    End If
   sWhere = IIf(Azione = TipoAzione.Modifica, "CodiceAnticipi='" & Replace(TxtCodice, "'", "''") & "'", "")
     If Response = vbYes Then
       If Not IsTableLocked("Anticipi") Then
            'LockTable ("Anticipi")
            If Azione = TipoAzione.Nuovo And Not GetADORecordset("Anticipi", "CodiceTribunale", "CodiceTribunale='" & cmbTribunale.Columns(1).value & "'  And CodiceAttivita='" & CmbAttivita.Columns(1).value & "'and CodiceAlternativo='" & CmbCodiceAlternativo.Columns(1).value & "'", g_Settings.DBConnection) Is Nothing Then
              MsgBox "Esiste già l'anticipo per " & cmbTribunale.Columns(0).value & ", " & CmbAttivita.Columns(0).value & " e " & CmbCodiceAlternativo.Columns(0).value, vbOKOnly + vbInformation, "Attenzione"
              Else
              Salvato = SalvaRecord(Me, Azione, "Anticipi", False, sWhere, True)
              
            End If
            'DelockTable ("Anticipi")
         Else
           MsgBox "Tabella Bloccata"
       End If
     End If
      If Salvato Then Visualizza (Azione = TipoAzione.Modifica)
  Else
   MsgBox "Azione non valida!", vbCritical + vbOKOnly
End If
If Salvato Then
  MsgBox "Anticipo salvato correttamente.", vbInformation + vbOKOnly
  CmdAnnulla_Click
End If


End Sub

Private Sub SettaCombo(cmb As ComboBox, code As String, Col As Collection)
Dim i As Integer
 For i = 1 To Col.Count
   If Col(i) = code Then Exit For
 Next i
 
 cmb.ListIndex = i - 1

End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub flex_BeforeSort(ByVal Col As Long, Order As Integer)
Call sortGrid(flex, Col, Order, 1, 1)
End Sub

Private Sub flex_DblClick()
Edit
End Sub

Private Sub Form_Load()
  
    
    
    PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", True, , "DescrizioneTribunale"
    
    
    
    PopolaTDBCombo CmbAttivita, "Attività", "Descrizione", "CodiceAttivita", True
    
    
    
    Query = "SELECT CodiceAnticipi as codice, Descrizione, PrezDepositoEuro as Deposito, PrezCompetenzeEuro as Competenze, PrezFormulaEuro as formula, PrezFotocopieEuro as Fotocopie, PrezCancelleriaEuro as Cancelleria, PrezMarcheEuro as Marche FROM Anticipi ORDER BY CodiceAnticipi"

    Visualizza False
    
End Sub
Public Sub Visualizza(storePosition As Boolean)
   Dim r As Long
    If storePosition Then r = flex.row
    

    AggiornaGriglia flex, Query
    FrmPrezzi.Visible = False
    
    If storePosition Then flex.row = r
End Sub
Public Sub Edit()
Dim r As Long
Dim rs As ADODB.Recordset
Dim code As String

Azione = TipoAzione.Modifica
r = flex.row
code = flex.TextMatrix(r, 1)
Set rs = GetADORecordset("Anticipi", "*", "CodiceAnticipi='" & code & "'", g_Settings.DBConnection)


  code = rs("CodiceAttivita")
 
 'CODICE ALTERNATICO
 LblCodAlternativo.Visible = (code <> "A")
 LblAutorita.Visible = (code <> "A")
 CmbCodiceAlternativo.Visible = (code <> "A")
 
 Select Case code
  Case "A"
            LblCodAlternativo.Caption = "XXX :"
            Dim rsFake As ADODB.Recordset
            Set rsFake = newAdoRs
            rsFake.Open "SELECT First('XXX'),'A' FROM Parametri", g_Settings.DBConnection
            Set CmbCodiceAlternativo.RowSource = rsFake
            
            
  Case "S"
            LblCodAlternativo.Caption = "Pignoramenti :"
            PopolaTDBCombo CmbCodiceAlternativo, "Pignoramenti", "Descrizione", "Codice", False, True


  Case "D"
            LblCodAlternativo.Caption = "Autorità :"
            PopolaTDBCombo CmbCodiceAlternativo, "Autorita", "Descrizione", "Codice", False, True
            

  Case "N"
            LblCodAlternativo.Caption = "Codice Atto :"
            PopolaTDBCombo CmbCodiceAlternativo, "TipoAtto", "Descrizione", "Codice", False, True
            
   Case Else
        
     CmbCodiceAlternativo.Clear
  End Select
  
 Caricacampi Me, rs, True
 FrmPrezzi.Visible = True

 rs.Close
 
 VisualizzaPrezzi (code)
 
End Sub
Public Sub VisualizzaPrezzi(code As String)
Dim rs As ADODB.Recordset
Dim i As Integer
If code = "" Then Exit Sub
 Set rs = GetADORecordset("Attività", "*", "CodiceAttivita='" & code & "'", g_Settings.DBConnection)
 For i = 1 To 6
   LblPrezzo(i).Visible = IIf(IsNull(rs(LblPrezzo(i).DataField)), False, (rs(LblPrezzo(i).DataField) = "S"))
   txtPrezzo(i).Visible = LblPrezzo(i).Visible
 Next i
 rs.Close
 
End Sub
 


Private Sub pulisci()
   ' Ripristino situazione TxtField
    pulisciCodice
    pulisciPrezzi
    
    ' Ripristino situazione ComboBox
    cmbTribunale.Text = ""
    CmbAttivita.Text = ""
    
    ' Ripristino situazione Btn
    CmdElimina.Enabled = False
    
    CmdAggiungi.Enabled = True
     
    FrmRicercaAnticipi.Enabled = True
End Sub

Private Sub pulisciPrezzi()
Dim i As Integer
For i = 1 To 6
  txtPrezzo(i).value = 0
Next i
    
End Sub

Private Sub pulisciCodice()
    TxtDescrizione.Text = ""
    TxtCodice.Text = ""
    LblAutorita.Caption = ""
End Sub

Public Sub GestioneCodiceAlternativo()

    LblCodAlternativo.Visible = True
    
    CmbCodiceAlternativo.Visible = True
    
    CmbCodiceAlternativo.Clear
    CmbCodiceAlternativo.Text = ""
    LblAutorita.Visible = True
    Dim code As String
     code = CmbAttivita.Columns(1).value
Select Case code

   Case "D"
      LblCodAlternativo.Caption = "Autorità :"
      PopolaTDBCombo CmbCodiceAlternativo, "Autorita", "Descrizione", "Codice", False, True
   Case "N"
      LblCodAlternativo.Caption = "Tipo Atto :"
      PopolaTDBCombo CmbCodiceAlternativo, "TipoAtto", "Descrizione", "Codice", False, True
   Case "S"
      LblCodAlternativo.Caption = "Pignoramenti :"
      PopolaTDBCombo CmbCodiceAlternativo, "Pignoramenti", "Descrizione", "Codice", False, True
   Case "A"
            LblCodAlternativo.Caption = "XXX :"
            Dim rsFake As ADODB.Recordset
            Set rsFake = newAdoRs
            rsFake.Open "SELECT First('XXX'),'A' FROM Parametri", g_Settings.DBConnection
            Set CmbCodiceAlternativo.RowSource = rsFake
           
           
    LblCodAlternativo.Visible = False
    CmbCodiceAlternativo.Visible = False
    
    LblAutorita.Visible = False
   Case Else
    MsgBox "Codice non previsto '" & code & "'. I codici validi sono:" & vbCrLf & _
           "A: Adempimenti;" & vbCrLf & _
           "D: Decreti Ingiuntivi;" & vbCrLf & _
           "N: Notifiche;" & vbCrLf & _
           "S: Sfratti;", vbCritical + vbOKOnly
           
End Select


End Sub

Private Sub txtPrezzo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")

End Sub
