VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmRicerca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ricerca"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14115
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   495
      Left            =   12720
      TabIndex        =   7
      Top             =   9720
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid flex 
      Height          =   7935
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   13935
      _cx             =   24580
      _cy             =   13996
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
      AllowUserResizing=   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraTop 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13755
      Begin VB.Frame fraAdempi 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   5760
         TabIndex        =   30
         Top             =   1080
         Width           =   6375
      End
      Begin VB.Frame fraUNEP 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   8280
         TabIndex        =   27
         Top             =   120
         Width           =   3855
         Begin VB.TextBox txtCrono 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Cronologico"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   3960
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbSiglaCh 
         Height          =   315
         Left            =   3120
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   360
         Width           =   1725
      End
      Begin VB.ComboBox cmbSigla 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   360
         Width           =   1725
      End
      Begin TDBDate6Ctl.TDBDate Da 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Calendar        =   "frmRicerca.frx":0000
         Caption         =   "frmRicerca.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRicerca.frx":0184
         Keys            =   "frmRicerca.frx":01A2
         Spin            =   "frmRicerca.frx":0200
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
      Begin VB.OptionButton opt 
         Caption         =   "Inevasi"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   700
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Evasi"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   700
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Tutto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   700
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdFiltra 
         Caption         =   "Filtra"
         Height          =   735
         Left            =   12240
         Picture         =   "frmRicerca.frx":0228
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxtRicCodAvv 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
      Begin TDBDate6Ctl.TDBDate A 
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Calendar        =   "frmRicerca.frx":0372
         Caption         =   "frmRicerca.frx":048A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRicerca.frx":04F6
         Keys            =   "frmRicerca.frx":0514
         Spin            =   "frmRicerca.frx":0572
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
         Left            =   4920
         TabIndex        =   32
         Tag             =   "necessario Tribunale"
         Top             =   360
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
         _PropDict       =   $"frmRicerca.frx":059A
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
         Left            =   4920
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Chiusura"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Inserimento"
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label LblRicCodAvv 
         Caption         =   "Cod. Cassetta :"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1410
      End
   End
   Begin VB.Frame fraAna 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   12795
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuova Cassetta"
         Height          =   375
         Left            =   11280
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdFiltraAna 
         Caption         =   "Filtra"
         Height          =   615
         Left            =   11280
         Picture         =   "frmRicerca.frx":0621
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxtUsufruente 
         Height          =   330
         Left            =   1785
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1005
         Width           =   3870
      End
      Begin VB.TextBox TxtRicNome 
         Height          =   330
         Left            =   1785
         MaxLength       =   40
         TabIndex        =   13
         Top             =   525
         Width           =   3870
      End
      Begin VB.TextBox TxtRicCodAvvInt 
         Height          =   330
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label LblUsufruente 
         Caption         =   "Usufruente :"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label LblRicNome 
         Caption         =   "Cognome e Nome :"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label LblRicCodAvvInt 
         Caption         =   "Cod. Cassetta :"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   165
         Width           =   1500
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuElimina 
         Caption         =   "Elimina"
      End
   End
End
Attribute VB_Name = "frmRicerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmCaller As IForm
Public Query As String
Public Tabella As String
Public Titolo As String
Public DefaultOrder As String
Public NCol As Integer
Public PosizioneCodice As Integer
Public tipo As String
Public Filtro As String
Public Location As Long
Dim qryExe As String

Private mFoundCode As String
Public isUnep As Boolean
Public Event AvvocatoSelezionato(codice As String)

Public Property Get FoundCode() As String
  FoundCode = mFoundCode
End Property





Private Sub cmbDate_Click()
  Dim data1 As Date
  Dim data2 As Date
  data2 = Date + 30
  data2 = LastDay(month(data2), year(data2))
  A = data2
 
Select Case cmbDate.ListIndex
  Case 0 'mese
    data1 = Date - 30
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 1 'trimestre
    data1 = Date - 90
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 2 'Semestre
    data1 = Date - 180
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 3 'Anno
    data1 = Date - 365
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 4 'Anno Completo
    Da = 1 & "/" & 1 & "/" & year(Date)
    A = 31 & "/" & 12 & "/" & year(Date)
  Case 5 'Tutto
    Da = ""
    A = ""
End Select


cmdFiltra_Click
End Sub

Private Sub cmbSigla_Click()
 cmdFiltra_Click
End Sub


Private Sub cmbSigla_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub cmbSiglaCh_Click()
cmdFiltra_Click
End Sub

Private Sub cmbSiglaCh_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
 AnagAvvocati.Azione = TipoAzione.Nuovo
 AnagAvvocati.Show
End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdFiltra_Click()
Dim Field As String
Field = "DataRegistrazione"
qryExe = Query
qryExe = qryExe & " WHERE 1=1 " & Filtro
    If txtCrono.Text <> "" Then
        qryExe = qryExe & " AND (Crono LIKE '%" & txtCrono.Text & "%')"
    End If
    If TxtRicCodAvv.Text <> "" Then
        qryExe = qryExe & " AND (CODAVV = '" & TxtRicCodAvv.Text & "')"
     End If
     If Tabella = "DEDUZIONI_UNEP" Then Field = "DataEvasionePratica"
     
    If IsDate(Da) Then
        qryExe = qryExe & " AND ( " & Field & " >= '" & Format(Da, "yyyymmdd") & "')"
    End If
    If IsDate(A) Then
        qryExe = qryExe & " AND ( " & Field & " <= '" & Format(A, "yyyymmdd") & "')"
    End If
    If opt(1) Then
       qryExe = qryExe & " AND ( CheckVisual = 'X')"
    End If
    If opt(2) Then
       qryExe = qryExe & " AND ( CheckVisual <> 'X')"
    End If
    If cmbSigla.ListIndex > 0 Then
       qryExe = qryExe & " AND ( SIGLA ='" & cmbSigla.list(cmbSigla.ListIndex) & "')"
    End If
    If cmbSiglaCh.ListIndex > 0 Then
       qryExe = qryExe & " AND ( SIGLACH ='" & cmbSiglaCh.list(cmbSiglaCh.ListIndex) & "')"
    End If
    Dim codTribunale
    If Not IsNull(cmbTribunale.SelectedItem) Then
        codTribunale = cmbTribunale.Columns(1).value
        If codTribunale <> "XXALLXX" Then qryExe = qryExe & " AND ( CodTribunaleApp ='" & codTribunale & "')"
    End If
  
    
    AggiornaGriglia
End Sub

Private Sub cmdFiltraAna_Click()
Screen.MousePointer = vbHourglass
qryExe = Query
qryExe = qryExe & " WHERE 1=1" & Filtro

    If TxtRicCodAvvInt.Text <> "" Then
        qryExe = qryExe & " AND(AnagraficaAvvocati.CODAVV  LIKE '" & TxtRicCodAvvInt.Text & "%')"
    End If
    
    If TxtRicNome.Text <> "" Then
        qryExe = qryExe & " AND(AnagraficaAvvocati.NOME Like '" & Replace(TxtRicNome.Text, "'", "''") & "%')"
    End If
    
    If TxtUsufruente.Text <> "" Then
        qryExe = "SELECT AnagraficaAvvocati.CODAVV as Codice, AnagraficaAvvocati.NOME as Nome,  AnagraficaAvvocati.Telef as Telefono,AnagraficaAvvocati.TelefCell as Cellulare,NumOrdinamento "
        qryExe = qryExe & "FROM AnagraficaAvvocati INNER JOIN Usufruenti ON AnagraficaAvvocati.CODAVV = Usufruenti.CODAVV"
        qryExe = qryExe & " WHERE (((Usufruenti.DescrizioneUsufr) Like '" & Replace(TxtUsufruente.Text, "'", "''") & "%'))"
    End If
    AggiornaGriglia
Screen.MousePointer = vbDefault
End Sub

Private Sub flex_AfterSort(ByVal Col As Long, Order As Integer)
ColoraAnnullati
If Col = 3 Then
  AggiornaGriglia
 Else
  sortGrid flex, Col, Order, 1, -1
End If
End Sub

Private Sub flex_DblClick()

Dim r As Long
r = flex.row

If Not frmCaller Is Nothing Then
  frmCaller.IsLoading = True
End If
If r < 1 Then Exit Sub
If tipo = "Anagrafica" Then
   If Not frmCaller Is Nothing Then
      If TypeOf frmCaller Is IAnagraficForm Then
        Dim iAnaForm As IAnagraficForm
        Set iAnaForm = frmCaller
        iAnaForm.RisultatoRicerca flex.TextMatrix(r, 1), TipoAzione.Nuovo
      End If
   End If
   mFoundCode = flex.TextMatrix(r, 1)
   RaiseEvent AvvocatoSelezionato(mFoundCode)
 Else
   If Not frmCaller Is Nothing Then
     
     
     frmCaller.Where = "IDCod= " & flex.TextMatrix(r, PosizioneCodice)
     frmCaller.RisRicerca
    
    
   End If
   mFoundCode = flex.TextMatrix(r, 1)
   RaiseEvent AvvocatoSelezionato(mFoundCode)
End If
 If Not frmCaller Is Nothing Then
    frmCaller.SetFocus
    
    frmCaller.IsLoading = False
 End If
  
'Unload Me
End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If flex.ColIndex("Nome") <> -1 Then
    If tipo = "Anagrafica" And Button = 2 And flex.TextMatrix(flex.row, flex.ColIndex("Nome")) = "" Then
      PopupMenu mnuContext
    End If
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Caption = Titolo
txtCrono.Text = ""
If tipo = "Anagrafica" Then
  Query = "SELECT CODAVV as Codice, NOME as Nome,Telef as Telefono, TelefCell as Cellulare,EMAIL as Mail,PEC, Mail2, NumOrdinamento FROM AnagraficaAvvocati "
  DefaultOrder = "order by AnagraficaAvvocati.NumOrdinamento"
  fraAna.Visible = True
  fraTop.Visible = False
  NCol = 7  'Numero di colonne da visualizzare
 Else
  fraAna.Visible = False
  fraTop.Visible = True
  PopolaCombo cmbSigla, "SELECT DISTINCT SIGLA as s FROM " & Tabella, "s", , , True
  PopolaCombo cmbSiglaCh, "SELECT DISTINCT SIGLACH as s FROM " & Tabella, "s", , , True
 
  PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", True
  
 
  cmbDate.AddItem "Mese"
  cmbDate.AddItem "Trimestre"
  cmbDate.AddItem "Semestre"
  cmbDate.AddItem "Anno"
  cmbDate.AddItem "Anno " & year(Date)
  cmbDate.AddItem "Tutto"
  cmbDate.ListIndex = 5
  fraUNEP.Visible = Tabella = "NOTIFICHE_UNEP" Or Tabella = "SFRATTI_UNEP"
  
  opt(0).Visible = Tabella <> "DEDUZIONI_UNEP"
  opt(1).Visible = Tabella <> "DEDUZIONI_UNEP"
  opt(2).Visible = Tabella <> "DEDUZIONI_UNEP"
  cmbSiglaCh.Visible = Tabella <> "DEDUZIONI_UNEP"
  Label2.Visible = Tabella <> "DEDUZIONI_UNEP"
End If
 qryExe = Query & " WHERE 1=1" & Filtro
 AggiornaGriglia
 Ridimensiona
If tipo = "Anagrafica" Then
  
Else
  flex.ColDataType(3) = flexDTDate
  
End If
Screen.MousePointer = vbDefault
End Sub
Public Sub Ridimensiona()
If Atap.ScaleHeight - Atap.Toolbar1.Height > 0 Then Me.Move Location, 0, Me.Width, Atap.ScaleHeight
If Me.Height - flex.Top - 600 > 0 Then flex.Height = Me.ScaleHeight - flex.Top - 600
cmdEsci.Top = flex.Top + flex.Height + 80
End Sub
Public Sub AggiornaGriglia()
Dim I As Integer
Dim rs As ADODB.Recordset
Set rs = newAdoRs
rs.Open qryExe & " " & DefaultOrder, g_Settings.DBConnection
Set flex.DataSource = rs
For I = NCol + 1 To flex.Cols - 1
 flex.ColHidden(I) = True
Next I

For I = 1 To flex.Cols - 1
 flex.ColWidth(I) = 20
Next I
flex.ColWidth(flex.ColIndex("Codice")) = 900
If tipo = "Anagrafica" Then
    flex.ColWidth(flex.ColIndex("Nome")) = 3800
    flex.ColWidth(flex.ColIndex("Telefono")) = 1500
    flex.ColWidth(flex.ColIndex("Cellulare")) = 1500
    flex.ColWidth(flex.ColIndex("Mail")) = 3200
    flex.ColWidth(flex.ColIndex("Pec")) = 3200
    flex.ColWidth(flex.ColIndex("Mail2")) = 3200
 Else
 
     flex.ColWidth(flex.ColIndex("Ev")) = 200
     flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1600
     flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1400
     flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1300
     
     flex.ColAlignment(flex.ColIndex("Data Registrazione")) = flexAlignCenterCenter
     flex.ColAlignment(flex.ColIndex("Sigla Inserimento")) = flexAlignCenterCenter
     flex.ColAlignment(flex.ColIndex("Sigla chiusura")) = flexAlignCenterCenter

    Select Case Tabella
      Case "ADEMPI"
       flex.ColWidth(flex.ColIndex("Attività")) = 7000
      Case "SFRATTI"
        flex.ColWidth(flex.ColIndex("Parte1")) = 3400
        flex.ColWidth(flex.ColIndex("Parte2")) = 3400
      Case "SFRATTI_UNEP"
      flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1500
        flex.ColWidth(flex.ColIndex("Parte1")) = 2800
        flex.ColWidth(flex.ColIndex("Parte2")) = 2800
        flex.ColWidth(flex.ColIndex("Crono")) = 2200
        flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1200
     flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1200
      Case "NOTIFICHE"
        flex.ColWidth(flex.ColIndex("Parte1")) = 3400
        flex.ColWidth(flex.ColIndex("Parte2")) = 3400
      Case "DecretiIngiuntivi"
        flex.ColWidth(flex.ColIndex("Ricorrente")) = 3400
        flex.ColWidth(flex.ColIndex("Debitore")) = 3400
      Case "NOTIFICHE_UNEP"
        flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1500
        flex.ColWidth(flex.ColIndex("Parte1")) = 2800
        flex.ColWidth(flex.ColIndex("Parte2")) = 2800
        flex.ColWidth(flex.ColIndex("Crono")) = 2200
        flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1200
        flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1200
     Case "DEDUZIONI_UNEP"
        flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1500
        flex.ColWidth(flex.ColIndex("Nota")) = 3000
        flex.ColAlignment(flex.ColIndex("Nota")) = flexAlignLeftCenter
        flex.ColFormat(flex.ColIndex("Importo")) = "0.00"
        flex.ColWidth(flex.ColIndex("Importo")) = 1500
        flex.ColWidth(flex.ColIndex("Ev")) = 300
    End Select
   
    
    
    
    
    'flex.ColWidth(flex.ColIndex("Parte1")) = 2000
    flex.ColWidth(0) = 200
    
    
 
End If
  ColoraAnnullati
End Sub

Private Sub ColoraAnnullati()
Dim I As Long
ColoraLiberi
ColoraEvasi
If flex.ColIndex("Annullo") = -1 Then Exit Sub
  For I = 1 To flex.Rows - 1
    
    If flex.TextMatrix(I, flex.ColIndex("Annullo")) = "A" Then
      flex.Cell(flexcpForeColor, I, 1, I, flex.Cols - 1) = &HC0C0C0
      flex.Cell(flexcpFontStrikethru, I, 1, I, flex.Cols - 1) = True
    End If
  Next I
 
End Sub

Private Sub ColoraEvasi()
Dim I As Long
If flex.ColIndex("Ev") = -1 Then Exit Sub
  For I = 1 To flex.Rows - 1
    
    If flex.TextMatrix(I, flex.ColIndex("Ev")) = "X" Then
      flex.Cell(flexcpForeColor, I, 1, I, flex.Cols - 1) = &HFF0000
      
    End If
  Next I
 
End Sub
Private Sub ColoraLiberi()
Dim I As Long
If flex.ColIndex("Telefono") = -1 Then Exit Sub
  For I = 1 To flex.Rows - 1
    
    If flex.TextMatrix(I, flex.ColIndex("Nome")) = "" Then
      flex.Cell(flexcpBackColor, I, 1, I, flex.Cols - 1) = &H80FF80
      
    End If
  Next I
 
End Sub


Private Sub mnuElimina_Click()
 Dim r As Long
 Dim codAvv As String
 codAvv = flex.TextMatrix(flex.row, 1)
 
 r = MsgBox("Sei sicuro di voler eliminare la cassetta " & codAvv)
 If r = vbOK Then
    g_Settings.DBConnection.Execute "DELETE * FROM AnagraficaAvvocati Where CodAvv='" & codAvv & "'"
    AggiornaGriglia
 End If
 
End Sub

Private Sub opt_Click(Index As Integer)
 cmdFiltra_Click
End Sub


Private Sub TxtRicCodAvv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdFiltra_Click
End Sub
