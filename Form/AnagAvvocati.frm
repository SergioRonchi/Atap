VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form AnagAvvocati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Anagrafica Avvocati"
   ClientHeight    =   7260
   ClientLeft      =   420
   ClientTop       =   600
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "E&sci"
      Height          =   450
      Left            =   6960
      TabIndex        =   45
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AnagAvvocati.frx":0000
            Key             =   "aperto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AnagAvvocati.frx":0452
            Key             =   "chiuso"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtCodiceAvvocatoInt 
      DataField       =   "codavv"
      Height          =   285
      Left            =   360
      MaxLength       =   10
      TabIndex        =   0
      Top             =   650
      Width           =   1305
   End
   Begin VB.Frame FrmUsufruenti 
      Caption         =   " Usufruenti "
      Height          =   2085
      Left            =   90
      TabIndex        =   37
      Top             =   4440
      Width           =   8115
      Begin VB.CommandButton CmdAggiungiUsuf 
         Caption         =   ">>"
         Height          =   240
         Left            =   3870
         TabIndex        =   16
         Top             =   270
         Width           =   405
      End
      Begin VB.CommandButton CmdEliminaUsuf 
         Caption         =   "<<"
         Height          =   240
         Left            =   3870
         TabIndex        =   17
         Top             =   585
         Width           =   405
      End
      Begin VB.ListBox LstUsufruenti 
         Height          =   1815
         ItemData        =   "AnagAvvocati.frx":08A4
         Left            =   4440
         List            =   "AnagAvvocati.frx":08A6
         TabIndex        =   18
         Top             =   120
         Width           =   3525
      End
      Begin VB.TextBox TxtNewUsufruente 
         Height          =   285
         Left            =   270
         MaxLength       =   30
         TabIndex        =   15
         Top             =   405
         Width           =   3510
      End
   End
   Begin VB.Frame FrmComandiAnag 
      Height          =   720
      Left            =   90
      TabIndex        =   35
      Top             =   6480
      Width           =   8130
      Begin VB.CommandButton cmdPrint 
         Height          =   465
         Left            =   5640
         Picture         =   "AnagAvvocati.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Stampa Schermata"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   450
         Left            =   4440
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame FrmAnagrafica 
      Height          =   4460
      Left            =   90
      TabIndex        =   19
      Top             =   0
      Width           =   8115
      Begin VB.TextBox txtOrdine 
         DataField       =   "NumOrdinamento"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Frame FrmNote 
         Height          =   1080
         Left            =   1200
         TabIndex        =   38
         Top             =   3240
         Width           =   5775
         Begin VB.TextBox TxtNote3 
            DataField       =   "NOTE3"
            Height          =   285
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   14
            Top             =   700
            Width           =   3870
         End
         Begin VB.TextBox TxtNote2 
            DataField       =   "NOTE2"
            Height          =   285
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   13
            Top             =   400
            Width           =   3870
         End
         Begin VB.TextBox TxtNote1 
            DataField       =   "NOTE1"
            Height          =   285
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   12
            Top             =   120
            Width           =   3870
         End
         Begin VB.Label LblNote 
            Caption         =   "Note :"
            Height          =   255
            Left            =   270
            TabIndex        =   39
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.TextBox TxtCognomeNome 
         DataField       =   "Nome"
         Height          =   285
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   1
         Top             =   650
         Width           =   4590
      End
      Begin VB.CheckBox ChkBlocco 
         Caption         =   "Libera Cassetta"
         DataField       =   "STAT"
         Height          =   795
         Left            =   6360
         Picture         =   "AnagAvvocati.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1680
      End
      Begin VB.CheckBox ChkFatturare 
         Caption         =   "Check1"
         DataField       =   "AFAT"
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   2400
         Width           =   240
      End
      Begin VB.TextBox TxtEMail 
         DataField       =   "EMAIL"
         Height          =   285
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2880
         Width           =   2925
      End
      Begin VB.TextBox TxtFax 
         DataField       =   "FAX"
         Height          =   285
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2880
         Width           =   1980
      End
      Begin VB.TextBox TxtCellulare 
         DataField       =   "TELEFCELL"
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2370
         Width           =   1980
      End
      Begin VB.TextBox TxtTelefono 
         DataField       =   "TELEF"
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2880
         Width           =   1980
      End
      Begin VB.TextBox TxtCodiceFiscale 
         DataField       =   "CFISC"
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   7
         Top             =   2370
         Width           =   1860
      End
      Begin VB.TextBox TxtPartitaIVA 
         DataField       =   "PIVA"
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   6
         Top             =   2370
         Width           =   1170
      End
      Begin VB.TextBox TxtCAP 
         DataField       =   "CAP"
         Height          =   285
         Left            =   240
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1800
         Width           =   720
      End
      Begin VB.TextBox TxtLocalità 
         DataField       =   "Locali"
         Height          =   285
         Left            =   1200
         MaxLength       =   35
         TabIndex        =   4
         Top             =   1800
         Width           =   3420
      End
      Begin VB.TextBox TxtIndirizzo 
         DataField       =   "Indiri"
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1200
         Width           =   3870
      End
      Begin TrueOleDBList80.TDBCombo CmbProvincia 
         DataField       =   "PROV"
         Height          =   315
         Left            =   5040
         TabIndex        =   43
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
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
         _PropDict       =   $"AnagAvvocati.frx":0E34
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
      Begin VB.Label LblAnagAnnullata 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Anagrafica Avvocato Annullata"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label LblFatturare 
         Caption         =   "Fatturare :"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label LblEMail 
         Caption         =   "E Mail :"
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   2655
         Width           =   675
      End
      Begin VB.Label LblFax 
         Caption         =   "Fax :"
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   2655
         Width           =   855
      End
      Begin VB.Label LblCellulare 
         Caption         =   "Cellulare :"
         Height          =   255
         Left            =   5040
         TabIndex        =   31
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label LblTelefono 
         Caption         =   "Telefono :"
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   2660
         Width           =   855
      End
      Begin VB.Label LblCodiceFiscale 
         Caption         =   "Codice Fiscale :"
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label LblPartitaIVA 
         Caption         =   "Partita IVA :"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label LblProvincia 
         Caption         =   "Provincia :"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LblCAP 
         Caption         =   "C.a.p. :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label LblLocalità 
         Caption         =   "Località :"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LblIndirizzo 
         Caption         =   "Indirizzo :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LblCognomeNome 
         Caption         =   "Cognome Nome :"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   410
         Width           =   1305
      End
      Begin VB.Label LblCodiceAvvocatoInt 
         Caption         =   "Cod. Cassetta :"
         Height          =   255
         Left            =   225
         TabIndex        =   22
         Top             =   410
         Width           =   1695
      End
   End
   Begin VB.Frame fraLibera 
      Height          =   5535
      Left            =   90
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   8130
      Begin VB.Label LblCassetta 
         Caption         =   "Cassetta Libera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   1920
         TabIndex        =   42
         Top             =   2760
         Width           =   4110
      End
   End
End
Attribute VB_Name = "AnagAvvocati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Posizione As String
Public Azione As TipoAzione
Dim PassaLoad As Boolean
Implements IAnagraficForm
Implements IForm
Private Sub ChkBlocco_Click()
Dim XXX As Boolean
Dim MSG_Avviso As String
Dim Response As Long
If ChkBlocco = 1 Then
  ChkBlocco.Picture = ImageList1.ListImages("aperto").Picture
  ChkBlocco.Caption = "Assegna Cassetta"
  FrmComandiAnag.Visible = False
 Else
  ChkBlocco.Picture = ImageList1.ListImages("chiuso").Picture
  ChkBlocco.Caption = "Libera Cassetta"
  FrmComandiAnag.Visible = True
End If

If Not PassaLoad Then
    LblAnagAnnullata.Visible = (ChkBlocco = 1)
    fraLibera.Visible = (ChkBlocco = 1)

    If ChkBlocco = 1 Then
     'Procedura di fine Rapporto
        MSG_Avviso = "Si è sicuri di voler eseguire la procedura di fine rapporto?"
        Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
        If Response = vbYes Then    ' User chose Yes.
            
            LblAnagAnnullata.Visible = True
            
            
            'procedura estratto conto NON UNEP
            '--------------------------------------------------------------------------------------------
            Riempi_PRT_EstrattoContoX "01/01/1900", Date, TxtCodiceAvvocatoInt.Text, 1, 1, 1, 1, "N", False
            If GetADORecordset("PrtEstrattoConto", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
               MsgBox "Nessun dato! Nessun Estratto conto per la cassetta " + TxtCodiceAvvocatoInt.Text, vbInformation, "Attenzione"
              
            Else
                    Call Stampa.gestioneReport("", "", 0, crptToWindow, "EstrattoConto.rpt", 1, "Tipo='LIQUIDAZIONE'")
                    MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
                    MSG_Avviso = MSG_Avviso & "Continuare?"
                    Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
                    If Response = vbYes Then    ' User chose Yes.
                        Riempi_PRT_Sospesi "01/01/1900", Date, TxtCodiceAvvocatoInt.Text, "NULL", "NULL", False, True
                        
                        If GetADORecordset("PrtSospesi", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
                          MsgBox "Nessun dato sospeso!.", vbInformation, "Attenzione"
                          Else
                            Call Stampa.gestioneReport("PrtSospesi", "", 0, crptToWindow, "SospesiLiquidazione.rpt", 1)
                        End If
                        MSG_Avviso = "Verificare il buon esito della stampa!" & vbCrLf
                        MSG_Avviso = "I dati verranno ora esportati e cancellati dal database." & vbCrLf
                        MSG_Avviso = MSG_Avviso & "Continuare?"
                            Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
                            If Response = vbNo Then    ' User chose no.
                                LblAnagAnnullata.Visible = False
                                Exit Sub
                            End If
        
                        XXX = Trasferisci(g_Settings.StoricoLiquidazioniPath & "\LIQ" & TxtCodiceAvvocatoInt.Text & "_" & Format(Date, "YYYYMMDD") & ".mdb", "19000101", Format(Date, "YYYYMMDD"), False, TxtCodiceAvvocatoInt.Text, "ADNS")
                    End If
              End If
              
            SvuotaTabellaSaldi TxtCodiceAvvocatoInt.Text
            '--------------------------------------------------------------------------------------------
              'procedura estratto conto  UNEP
            '--------------------------------------------------------------------------------------------
            Riempi_PRT_EstrattoContoX "01/01/1900", Date, TxtCodiceAvvocatoInt.Text, 1, 1, 1, 1, "N", True
            If GetADORecordset("PrtEstrattoContoUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
               MsgBox "Nessun dato! Nessun Estratto UNEP conto per la cassetta " + TxtCodiceAvvocatoInt.Text, vbInformation, "Attenzione"
              
            Else
                    Call Stampa.gestioneReport("", "", 0, crptToWindow, "EstrattoContoUNEP.rpt", 1, "Tipo='LIQUIDAZIONE'")
                    MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
                    MSG_Avviso = MSG_Avviso & "Continuare?"
                    Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
                    If Response = vbYes Then    ' User chose Yes.
                        Riempi_PRT_Sospesi "01/01/1900", Date, TxtCodiceAvvocatoInt.Text, "NULL", "NULL", True, True
                        
                        If GetADORecordset("PrtSospesiUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
                          MsgBox "Nessun dato sospeso!.", vbInformation, "Attenzione"
                          Else
                            Call Stampa.gestioneReport("PrtSospesiUNEP", "", 0, crptToWindow, "SospesiLiquidazioneUNEP.rpt", 1)
                        End If
                        MSG_Avviso = "Verificare il buon esito della stampa!" & vbCrLf
                        MSG_Avviso = "I dati verranno ora esportati e cancellati dal database." & vbCrLf
                        MSG_Avviso = MSG_Avviso & "Continuare?"
                            Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
                            If Response = vbNo Then    ' User chose no.
                                LblAnagAnnullata.Visible = False
                                Exit Sub
                            End If
        
                        XXX = Trasferisci(g_Settings.StoricoLiquidazioniPath & "\LIQUNEP" & TxtCodiceAvvocatoInt.Text & "_" & Format(Date, "YYYYMMDD") & ".mdb", "19000101", Format(Date, "YYYYMMDD"), True, TxtCodiceAvvocatoInt.Text, "ADNS")
                    End If
              End If
              
            SvuotaTabellaSaldiUNEP TxtCodiceAvvocatoInt.Text
            '--------------------------------------------------------------------------------------------
            
            LiberaCasella TxtCodiceAvvocatoInt.Text
            PulisciCampi
            

           Else
            ChkBlocco = 0
        End If
      Else
      
    End If
End If
End Sub





Private Sub CmdAggiungiUsuf_Click()
    If Trim(TxtNewUsufruente.Text) <> "" Then
        LstUsufruenti.AddItem TxtNewUsufruente.Text
        TxtNewUsufruente.Text = ""
    End If
End Sub

Private Sub CmdEliminaUsuf_Click()
    
If LstUsufruenti.ListIndex >= 0 Then
    LstUsufruenti.RemoveItem LstUsufruenti.ListIndex
End If
    
End Sub

Private Sub cmdEsci_Click()
Atap.mnuAnagAvvoc_Click

Unload Me
End Sub

Private Sub cmdPrint_Click()
 PrintForm
End Sub


Private Sub CmdSalva_Click()

Dim Response As Variant
On Error GoTo FINE
g_Settings.DBConnection.BeginTrans

If TxtCognomeNome.Text = "" Then
    MsgBox "Il Cognome Nome deve essere per forza inserito!", vbInformation, "Attenzione"
    TxtCognomeNome.SetFocus
    Exit Sub
End If

If Trim(TxtCodiceAvvocatoInt.Text) = "" Then
    MsgBox "Il codice avvocato deve essere per forza inserito!", vbInformation, "Attenzione"
    TxtCodiceAvvocatoInt.SetFocus
    Exit Sub
End If

If ChkFatturare.value = vbChecked Then
    If Trim(TxtPartitaIVA.Text) = "" Or val(TxtPartitaIVA.Text) = 0 Then
            MsgBox "Inserire la partita iva!", vbInformation, "Attenzione"
            TxtPartitaIVA.SetFocus
            Exit Sub
    End If
    
    If Not CheckCodFiscPIva(TxtPartitaIVA.Text) Then
            MsgBox "La partita IVA non è corretta!", vbInformation, "Attenzione"
            TxtPartitaIVA.SetFocus
            Exit Sub
    End If
    
End If

If Trim(TxtPartitaIVA.Text) <> "" Then
    If Not CheckCodFiscPIva(TxtPartitaIVA.Text) Then
        MsgBox "La Partita Iva non è corretta!", vbInformation, "Attenzione"
        TxtPartitaIVA.SetFocus
        Exit Sub
    End If
End If

If Azione = TipoAzione.Modifica Then
    'Sto Modificando la mia anagrafica
    Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
        If Not IsTableLocked("AnagraficaAvvocati") Then
            'LockTable ("AnagraficaAvvocati")
            If SalvaRecord(Me, Azione, "AnagraficaAvvocati", False, "CODAVV='" & TxtCodiceAvvocatoInt & "'") Then
              salvaUsufruenti (TxtCodiceAvvocatoInt)
              If GetADORecordset("SALDI", "CODICE", "CODICE='" & TxtCodiceAvvocatoInt & "'", g_Settings.DBConnection) Is Nothing Then
                g_Settings.DBConnection.Execute "INSERT INTO SALDI (CODICE,Chiusura,NumOrdinamento) VALUES ('" & TxtCodiceAvvocatoInt & "','" & Format(Date, "YYYYMMDD") & "','" & txtOrdine & "')"
              End If
              If GetADORecordset("SALDIUNEP", "CODICE", "CODICE='" & TxtCodiceAvvocatoInt & "'", g_Settings.DBConnection) Is Nothing Then
                g_Settings.DBConnection.Execute "INSERT INTO SALDIUNEP (CODICE,Chiusura,NumOrdinamento) VALUES ('" & TxtCodiceAvvocatoInt & "','" & Format(Date, "YYYYMMDD") & "','" & txtOrdine & "')"
              End If
            End If
            'DelockTable ("AnagraficaAvvocati")
          Else
           MsgBox "Tabella Bloccata"
        End If
    End If
    'Call PulisciCampi
Else
    'Sto Aggiungendo un record alla mia anagrafica
    Response = MsgBox("Vuoi salvare i dati inseriti?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    
    If Response = vbYes Then    ' User chose Yes.
        If Not IsTableLocked("AnagraficaAvvocati") Then
            'LockTable "AnagraficaAvvocati"
            If SalvaRecord(Me, Azione, "AnagraficaAvvocati", False) Then
              salvaUsufruenti (TxtCodiceAvvocatoInt)
              g_Settings.DBConnection.Execute "INSERT INTO SALDI (CODICE,Chiusura,NumOrdinamento) VALUES ('" & TxtCodiceAvvocatoInt & "','" & Format(Date, "YYYYMMDD") & "','" & txtOrdine & "')"
              g_Settings.DBConnection.Execute "INSERT INTO SALDIUNEP (CODICE,Chiusura,NumOrdinamento) VALUES ('" & TxtCodiceAvvocatoInt & "','" & Format(Date, "YYYYMMDD") & "','" & txtOrdine & "')"
            End If
            'DelockTable "AnagraficaAvvocati"
          Else
           MsgBox "Tabella Bloccata"
       End If
    End If
    'Call PulisciCampi
    CmdSalva.Caption = "&Aggiungi"
'    LstUsufruenti.Enabled = True
    CmdSalva.Enabled = False
    ChkBlocco.Visible = True
End If
FrmUsufruenti.Enabled = True
TxtNewUsufruente.Enabled = False
CmdAggiungiUsuf.Enabled = False
CmdEliminaUsuf.Enabled = False
g_Settings.DBConnection.CommitTrans
Unload Me
Exit Sub

FINE:
MsgBox err.Description & vbCrLf & "Salvataggio fallito"
g_Settings.DBConnection.RollbackTrans


End Sub
Private Function CalcolaOrdinamento() As Long
'Trova probabile posizione nell'ordinamento
        Dim n As Long, N2 As Long, N1 As Long, MyNum As Long
        Dim MyCod As String
        Dim C525 As String
        Dim C100 As String
        C525 = val(Mid(TxtCodiceAvvocatoInt, 5))
        C100 = val(TxtCodiceAvvocatoInt)
        MyCod = IIf(Left(TxtCodiceAvvocatoInt, 3) = "525", "Z" & String(6 - Len(C525), "0") & Mid(TxtCodiceAvvocatoInt, 5), "A" & String(6 - Len(C100), "0") & TxtCodiceAvvocatoInt)
          n = GetADOValue("AnagraficaAvvocati", "NumOrdinamento", "IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv)='" & MyCod & "' ORDER BY IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv) DESC", g_Settings.DBConnection, True)
          N1 = GetADOValue("AnagraficaAvvocati", "NumOrdinamento", "IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv)<'" & MyCod & "' ORDER BY IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv) DESC", g_Settings.DBConnection, True)
          N2 = GetADOValue("AnagraficaAvvocati", "NumOrdinamento", "IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv)>'" & MyCod & "' ORDER BY IIf(Left(CodAvv, 3) = '525', 'Z' & String(6 - Len(Val(Mid(CodAvv, 5))), '0') & Mid(CodAvv, 5), 'A' & String(6 - Len(Val(CodAvv)), '0') & CodAvv)", g_Settings.DBConnection, True)
          If N2 = 0 Then N2 = N1 + 100
          CalcolaOrdinamento = (N2 + N1) / 2
End Function
Private Sub salvaUsufruenti(cod As String)
 Dim i As Integer
 g_Settings.DBConnection.Execute "DELETE * FROM USUFRUENTI WHERE CODAVV='" & cod & "'"
 For i = 0 To LstUsufruenti.ListCount - 1
   g_Settings.DBConnection.Execute "INSERT INTO USUFRUENTI(CODAVV,DescrizioneUsufr) VALUES ('" & cod & "','" & Replace(LstUsufruenti.List(i), "'", "''") & "')"
 Next i
 
End Sub
Private Sub Form_Load()
    PassaLoad = True
    Me.Move 0, 0
    Call PopolaTDBCombo(CmbProvincia, "Provincie", "CodiceProvincia", "CodiceProvincia", , , "CodiceProvincia")
    CmdSalva.Caption = IIf(Azione = TipoAzione.Nuovo, "&Salva", "&Modifica")
    TxtCodiceAvvocatoInt.Enabled = (Azione = TipoAzione.Nuovo)
    ChkBlocco.Visible = (Azione <> TipoAzione.Nuovo)
    
    
    PassaLoad = False
End Sub




Public Sub PulisciCampi()
       TxtCAP.Text = ""
        TxtCellulare.Text = ""
        TxtCodiceFiscale.Text = ""
        TxtCognomeNome.Text = ""
        TxtEMail.Text = ""
        TxtFax.Text = ""
        TxtIndirizzo.Text = ""
        TxtLocalità.Text = ""
        TxtNote1.Text = ""
        TxtNote2.Text = ""
        TxtNote3.Text = ""
        TxtPartitaIVA.Text = ""
        CmbProvincia.Text = ""
        TxtTelefono.Text = ""
        ChkFatturare.value = Unchecked
        LstUsufruenti.Clear
        AbilitaCampi True
        'NascondiCampi True
      
        
End Sub


Public Function AbilitaCampi(SiNO As Boolean)
        TxtCAP.Enabled = SiNO
        TxtCellulare.Enabled = SiNO
        TxtCodiceFiscale.Enabled = SiNO
        TxtCodiceAvvocatoInt.Enabled = SiNO
   '     TxtCodiceAvvocatoUsuf.Enabled = SiNO
'        TxtDescrUsuf.Enabled = SiNO
        TxtCognomeNome.Enabled = SiNO
        TxtEMail.Enabled = SiNO
        TxtFax.Enabled = SiNO
        TxtIndirizzo.Enabled = SiNO
        TxtLocalità.Enabled = SiNO
        TxtNote1.Enabled = SiNO
        TxtNote2.Enabled = SiNO
        TxtNote3.Enabled = SiNO
        TxtPartitaIVA.Enabled = SiNO
        CmbProvincia.Enabled = SiNO
        TxtTelefono.Enabled = SiNO
        ChkFatturare.Enabled = SiNO
'        CmbZona.Enabled = SiNO
  '      CmbTribunale.Enabled = SiNO
End Function

Public Function NascondiCampi(SiNO As Boolean)
        lblCAP.Visible = SiNO
        LblCellulare.Visible = SiNO
        LblCodiceFiscale.Visible = SiNO
        LblCognomeNome.Visible = SiNO
        LblEMail.Visible = SiNO
        LblFax.Visible = SiNO
        LblIndirizzo.Visible = SiNO
        LblLocalità.Visible = SiNO
        LblPartitaIVA.Visible = SiNO
        LblProvincia.Visible = SiNO
        LblTelefono.Visible = SiNO
        LblFatturare.Visible = SiNO
        TxtCAP.Visible = SiNO
        TxtCellulare.Visible = SiNO
        TxtCodiceFiscale.Visible = SiNO
        TxtCognomeNome.Visible = SiNO
        TxtEMail.Visible = SiNO
        TxtFax.Visible = SiNO
        TxtIndirizzo.Visible = SiNO
        TxtLocalità.Visible = SiNO
        FrmNote.Visible = SiNO
        FrmUsufruenti.Visible = SiNO
        TxtPartitaIVA.Visible = SiNO
        CmbProvincia.Visible = SiNO
        TxtTelefono.Visible = SiNO
        ChkFatturare.Visible = SiNO
End Function



Public Sub RiempiLstUsufruenti()
Dim SQL As String
Dim rs As ADODB.Recordset

Set rs = newAdoRs
SQL = "SELECT * FROM USUFRUENTI WHERE CODAVV='" & TxtCodiceAvvocatoInt & "'"

rs.Open SQL, g_Settings.DBConnection

While Not rs.EOF
  LstUsufruenti.AddItem rs!DescrizioneUsufr
  rs.MoveNext
Wend
rs.Close

End Sub


Private Property Let IForm_IsLoading(RHS As Boolean)
 PassaLoad = RHS
End Property

Private Property Get IForm_IsLoading() As Boolean
 IForm_IsLoading = PassaLoad
End Property

Private Sub IForm_RisRicerca()

End Sub

Private Sub IForm_SetFocus()
 Me.SetFocus
End Sub

Private Property Let IForm_Where(RHS As String)

End Property

Private Sub TxtCodiceAvvocatoInt_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
 KeyAscii = KeyAscii - 32
End If

End Sub

Private Sub TxtCodiceAvvocatoInt_LostFocus()
txtOrdine = CalcolaOrdinamento()
'Test sull'esistenza del codice cassetta
If Azione = TipoAzione.Nuovo And Not GetADORecordset("AnagraficaAvvocati", "CODAVV", "CODAVV='" & TxtCodiceAvvocatoInt.Text & "'", g_Settings.DBConnection) Is Nothing Then
        MsgBox "Codice Cassetta Esistente!", vbInformation, "Attenzione"
        TxtCodiceAvvocatoInt.SetFocus
        SendKeys "{Home}+{End}"
End If

End Sub


Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocatoInt.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(sCodAvv As String, oAzione As TipoAzione)

Dim SQL As String
Dim rs As ADODB.Recordset
Azione = TipoAzione.Modifica
If Not FindForm("AnagAvvocati") Then Load AnagAvvocati
PassaLoad = True
Set rs = newAdoRs
SQL = "SELECT * FROM ANAGRAFICAAVVOCATI WHERE CODAVV='" & sCodAvv & "'"
rs.Open SQL, g_Settings.DBConnection
Caricacampi Me, rs
RiempiLstUsufruenti

    'Call RiempiCampi
    fraLibera.Visible = (rs!STAT = "A")
    FrmComandiAnag.Visible = Not (rs!STAT = "A")
    fraLibera.ZOrder
    LblAnagAnnullata.Visible = -ChkBlocco.value
PassaLoad = False
End Sub

Private Sub IAnagraficForm_SelectCodiceAvvocato()
 TxtCodiceAvvocatoInt.SetFocus
 SendKeys "{Home}+{End}"
End Sub

