VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmDecretiIngiuntivi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Decreti Ingiuntivi"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Tag             =   "NumeroDecreto"
   Begin VB.Frame fraMain 
      Height          =   5490
      Left            =   20
      TabIndex        =   34
      Top             =   720
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   4080
         Top             =   2160
      End
      Begin VB.TextBox txtQtaCopie 
         DataField       =   "QtaCopie"
         Height          =   285
         Left            =   8160
         TabIndex        =   82
         Tag             =   "necessario Copie"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtSigla 
         DataField       =   "SIGLA"
         Height          =   285
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "necessario Sigla Inserimento"
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame fraFormula 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   6285
         TabIndex        =   77
         Top             =   2250
         Visible         =   0   'False
         Width           =   2535
         Begin TDBNumber6Ctl.TDBNumber txtFormula 
            DataField       =   "ImpFormulaE"
            Height          =   285
            Left            =   1160
            TabIndex        =   22
            Top             =   0
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Calculator      =   "DecretiIngiuntivi.frx":0000
            Caption         =   "DecretiIngiuntivi.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":008C
            Keys            =   "DecretiIngiuntivi.frx":00AA
            Spin            =   "DecretiIngiuntivi.frx":00F4
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.Label Label6 
            Caption         =   "-"
            Height          =   255
            Left            =   1035
            TabIndex        =   79
            Top             =   0
            Width           =   135
         End
         Begin VB.Label LblFormula 
            Caption         =   "Formula : "
            Height          =   255
            Left            =   0
            TabIndex        =   78
            Top             =   20
            Width           =   1455
         End
      End
      Begin VB.Frame fraEsenti 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   660
         Left            =   240
         TabIndex        =   68
         Top             =   2320
         Width           =   4335
         Begin TDBNumber6Ctl.TDBNumber txtSpese 
            DataField       =   "ImpMarcheE"
            Height          =   285
            Index           =   2
            Left            =   2040
            TabIndex        =   15
            Top             =   0
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Calculator      =   "DecretiIngiuntivi.frx":011C
            Caption         =   "DecretiIngiuntivi.frx":013C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":01A8
            Keys            =   "DecretiIngiuntivi.frx":01C6
            Spin            =   "DecretiIngiuntivi.frx":0210
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber txtSpese 
            DataField       =   "ImpCopieE"
            Height          =   285
            Index           =   3
            Left            =   2040
            TabIndex        =   17
            Top             =   315
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Calculator      =   "DecretiIngiuntivi.frx":0238
            Caption         =   "DecretiIngiuntivi.frx":0258
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":02C4
            Keys            =   "DecretiIngiuntivi.frx":02E2
            Spin            =   "DecretiIngiuntivi.frx":032C
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber TxtQta 
            DataField       =   "QtaMarche"
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Calculator      =   "DecretiIngiuntivi.frx":0354
            Caption         =   "DecretiIngiuntivi.frx":0374
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":03E0
            Keys            =   "DecretiIngiuntivi.frx":03FE
            Spin            =   "DecretiIngiuntivi.frx":0448
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   ","
            DisplayFormat   =   "##0;;Null"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   1
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   2000
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber TxtQta 
            DataField       =   "QtaDirittiCancelleria"
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   16
            Top             =   320
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            Calculator      =   "DecretiIngiuntivi.frx":0470
            Caption         =   "DecretiIngiuntivi.frx":0490
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":04FC
            Keys            =   "DecretiIngiuntivi.frx":051A
            Spin            =   "DecretiIngiuntivi.frx":0564
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   ","
            DisplayFormat   =   "##0;;Null"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   1
            ForeColor       =   -2147483640
            Format          =   "##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   2000
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ""
            ShowContextMenu =   1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.Label LblDirittiDiCancelleria 
            Caption         =   "Diritti di Cancelleria:"
            Height          =   420
            Left            =   0
            TabIndex        =   70
            Top             =   240
            Width           =   960
         End
         Begin VB.Label LblMarche 
            Caption         =   "Marche :"
            Height          =   240
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   4320
         Width           =   4935
         Begin VB.TextBox txtSiglaCH 
            DataField       =   "SIGLACH"
            Height          =   285
            Left            =   4080
            MaxLength       =   3
            TabIndex        =   27
            Tag             =   "Sigla Chiusura"
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox ChkAnnullo 
            Caption         =   "Check1"
            DataField       =   "Annullo"
            Height          =   240
            Left            =   1440
            TabIndex        =   28
            Tag             =   "PULISCI"
            Top             =   645
            Width           =   240
         End
         Begin VB.CheckBox chkEvadi 
            Caption         =   "Check1"
            DataField       =   "CheckVisual"
            Height          =   240
            Left            =   3210
            TabIndex        =   26
            ToolTipText     =   "Evadi"
            Top             =   15
            Width           =   240
         End
         Begin TDBDate6Ctl.TDBDate txtDataEvaso 
            DataField       =   "DataEvasionePratica"
            Height          =   255
            Left            =   1410
            TabIndex        =   25
            Top             =   15
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "DecretiIngiuntivi.frx":058C
            Caption         =   "DecretiIngiuntivi.frx":06A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "DecretiIngiuntivi.frx":0710
            Keys            =   "DecretiIngiuntivi.frx":072E
            Spin            =   "DecretiIngiuntivi.frx":078C
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
         Begin VB.Label Label7 
            Caption         =   "Sigla : "
            Height          =   255
            Left            =   3480
            TabIndex        =   81
            Top             =   720
            Width           =   510
         End
         Begin VB.Label LblAvvDecAnn 
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
            Left            =   2040
            TabIndex        =   67
            Top             =   645
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label LblAnnullo 
            Caption         =   "Annulla decreto : "
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   645
            Width           =   1320
         End
         Begin VB.Label LblDescrEvaso 
            Caption         =   "Decreto evaso in data :"
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
            Left            =   1320
            TabIndex        =   65
            Tag             =   "PULISCI"
            Top             =   255
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label LblDataEvaso 
            Caption         =   "Data Evasione : "
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Frame fraMaschera 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   60
         Top             =   120
         Width           =   1095
         Begin VB.Label LblNumeroAtto 
            DataField       =   "NumeroDecreto"
            Height          =   255
            Left            =   480
            TabIndex        =   62
            Tag             =   "PULISCI"
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label LblDecreto 
            Caption         =   "Numero Decreto: "
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   75
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.TextBox TxtNumeroRuolo 
         DataField       =   "NumeroRuolo"
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   7
         Top             =   840
         Width           =   960
      End
      Begin VB.Frame FrmCommento 
         Caption         =   " Commento "
         Height          =   1515
         Left            =   5100
         TabIndex        =   58
         Top             =   3480
         Width           =   4635
         Begin VB.TextBox txtCommento 
            DataField       =   "Commento"
            Height          =   1155
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   24
            Top             =   240
            Width           =   4395
         End
      End
      Begin VB.CheckBox ChkFormulaEsec 
         Caption         =   "Check1"
         DataField       =   "FormulaEsec"
         Height          =   195
         Left            =   5880
         TabIndex        =   21
         Tag             =   "PULISCI"
         Top             =   2280
         Width           =   195
      End
      Begin VB.CheckBox ChkEsenzione 
         Caption         =   "Check1"
         DataField       =   "Esenzione"
         Height          =   195
         Left            =   2040
         TabIndex        =   13
         Tag             =   "PULISCI"
         Top             =   2010
         Width           =   285
      End
      Begin VB.TextBox TxtParte1 
         DataField       =   "Ricorrente"
         Height          =   285
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   9
         Tag             =   "necessario Ricorrente"
         Top             =   1155
         Width           =   2925
      End
      Begin VB.TextBox TxtDescrSpeseVarie 
         DataField       =   "DesrSpese"
         Height          =   285
         Left            =   1485
         MaxLength       =   35
         TabIndex        =   23
         Top             =   3720
         Width           =   2925
      End
      Begin VB.TextBox TxtNrIng 
         DataField       =   "NumeroIngiunzione"
         Height          =   285
         Left            =   5085
         MaxLength       =   6
         TabIndex        =   8
         Top             =   840
         Width           =   720
      End
      Begin VB.TextBox TxtParte2 
         DataField       =   "Debitore"
         Height          =   285
         Left            =   6210
         MaxLength       =   35
         TabIndex        =   10
         Tag             =   "necessario Debitore"
         Top             =   1155
         Width           =   3285
      End
      Begin TDBDate6Ctl.TDBDate txtDataReg 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Tag             =   "necessario Data Registrazione"
         Top             =   120
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calendar        =   "DecretiIngiuntivi.frx":07B4
         Caption         =   "DecretiIngiuntivi.frx":08CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0938
         Keys            =   "DecretiIngiuntivi.frx":0956
         Spin            =   "DecretiIngiuntivi.frx":09B4
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
         Left            =   7440
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "DecretiIngiuntivi.frx":09DC
         Caption         =   "DecretiIngiuntivi.frx":09FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0A68
         Keys            =   "DecretiIngiuntivi.frx":0A86
         Spin            =   "DecretiIngiuntivi.frx":0AD0
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
         DataField       =   "ImpFotocopieE"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "DecretiIngiuntivi.frx":0AF8
         Caption         =   "DecretiIngiuntivi.frx":0B18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0B84
         Keys            =   "DecretiIngiuntivi.frx":0BA2
         Spin            =   "DecretiIngiuntivi.frx":0BEC
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
         Left            =   7440
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "DecretiIngiuntivi.frx":0C14
         Caption         =   "DecretiIngiuntivi.frx":0C34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0CA0
         Keys            =   "DecretiIngiuntivi.frx":0CBE
         Spin            =   "DecretiIngiuntivi.frx":0D08
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
         Index           =   4
         Left            =   2280
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "DecretiIngiuntivi.frx":0D30
         Caption         =   "DecretiIngiuntivi.frx":0D50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0DBC
         Keys            =   "DecretiIngiuntivi.frx":0DDA
         Spin            =   "DecretiIngiuntivi.frx":0E24
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
      Begin TDBNumber6Ctl.TDBNumber TxtQta 
         DataField       =   "QtaFotocopie"
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         Calculator      =   "DecretiIngiuntivi.frx":0E4C
         Caption         =   "DecretiIngiuntivi.frx":0E6C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DecretiIngiuntivi.frx":0ED8
         Keys            =   "DecretiIngiuntivi.frx":0EF6
         Spin            =   "DecretiIngiuntivi.frx":0F40
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "##0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "##0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2000
         MinValue        =   0
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
      Begin TrueOleDBList80.TDBCombo cmbTribunale 
         DataField       =   "CodTribunaleApp"
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Tag             =   "necessario Tribunale"
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
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
         MatchEntryTimeout=   0
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
         _PropDict       =   $"DecretiIngiuntivi.frx":0F68
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
      Begin TrueOleDBList80.TDBCombo cmbAutorita 
         DataField       =   "CodAutorita"
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Tag             =   "necessario Autorità"
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
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
         MatchEntryTimeout=   0
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
         _PropDict       =   $"DecretiIngiuntivi.frx":0FEF
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
      Begin VB.Label Label3 
         Caption         =   "Sigla : "
         Height          =   255
         Left            =   2880
         TabIndex        =   80
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         Height          =   255
         Left            =   7320
         TabIndex        =   76
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "-"
         Height          =   255
         Left            =   7320
         TabIndex        =   75
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "+"
         Height          =   255
         Left            =   7320
         TabIndex        =   74
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label lblTotSpese 
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
         Index           =   1
         Left            =   7440
         TabIndex        =   73
         Tag             =   "PULISCI"
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Spese:"
         Height          =   255
         Left            =   6285
         TabIndex        =   72
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   3600
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label lblTotSpese 
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
         Index           =   0
         Left            =   2280
         TabIndex        =   71
         Tag             =   "PULISCI"
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Line Line1 
         X1              =   7320
         X2              =   8880
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblNumeroRuolo 
         Caption         =   "R.G."
         Height          =   255
         Left            =   225
         TabIndex        =   59
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label LblNumCopie 
         Caption         =   "Copie :"
         Height          =   240
         Left            =   7560
         TabIndex        =   57
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblCompetenze 
         Caption         =   "Competenze : "
         Height          =   255
         Left            =   6285
         TabIndex        =   56
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label LblFormulaEsec 
         Caption         =   "Formula Esec.  : "
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label LblFotocopie 
         Caption         =   "Fotocopie :"
         Height          =   240
         Left            =   240
         TabIndex        =   54
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label LblCosto 
         Caption         =   "Costo unitario"
         Height          =   330
         Left            =   2250
         TabIndex        =   53
         Top             =   1455
         Width           =   1140
      End
      Begin VB.Label LblQta 
         Caption         =   "Q.tà"
         Height          =   285
         Left            =   1530
         TabIndex        =   52
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label LblParte2 
         Caption         =   "Debitore : "
         Height          =   255
         Left            =   5085
         TabIndex        =   51
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label LblDataDep 
         Caption         =   "Data Deposito : "
         Height          =   255
         Left            =   225
         TabIndex        =   50
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label LblParte1 
         Caption         =   "Ricorrente : "
         Height          =   255
         Left            =   225
         TabIndex        =   49
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label LblDeposito 
         Caption         =   "Deposito : "
         Height          =   255
         Left            =   6285
         TabIndex        =   48
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label LblSpese 
         Caption         =   "Spese Cancelleria : "
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3030
         Width           =   1365
      End
      Begin VB.Label LblSaldo 
         Caption         =   "Saldo : "
         Height          =   255
         Left            =   6285
         TabIndex        =   46
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label LblDescrSpeseCancelleria 
         Caption         =   "Desc. Spese : "
         Height          =   255
         Left            =   270
         TabIndex        =   45
         Top             =   3720
         Width           =   1095
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
         Left            =   7440
         TabIndex        =   44
         Top             =   3240
         Width           =   1230
      End
      Begin VB.Label LblNrIng 
         Caption         =   "Numero Ingiunzione : "
         Height          =   255
         Left            =   3330
         TabIndex        =   42
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label LblAutorita 
         Caption         =   "Autorità : "
         Height          =   255
         Left            =   225
         TabIndex        =   39
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label LblEsenzione 
         Caption         =   "Esenzione (S)  x G. P.   : "
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2010
         Width           =   1770
      End
      Begin VB.Label LblTribunale 
         Caption         =   "Tribunale :"
         Height          =   255
         Left            =   5040
         TabIndex        =   36
         Top             =   480
         Width           =   825
      End
      Begin VB.Label LblDescrizioneAutorita 
         Height          =   240
         Left            =   3360
         TabIndex        =   35
         Tag             =   "PULISCI"
         Top             =   600
         Width           =   2220
      End
   End
   Begin VB.Frame fraComandi 
      Height          =   660
      Left            =   0
      TabIndex        =   33
      Top             =   6240
      Width           =   9855
      Begin VB.CommandButton cmdPrint 
         Height          =   500
         Left            =   4080
         Picture         =   "DecretiIngiuntivi.frx":1076
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Stampa Schermata"
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   500
         Left            =   6000
         TabIndex        =   29
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdAnnulla 
         Caption         =   "Esci"
         Height          =   500
         Left            =   7920
         TabIndex        =   43
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   500
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   1860
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "Ri&cerca Decreti"
         Height          =   500
         Left            =   2160
         TabIndex        =   40
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.Frame FraTop 
      Height          =   645
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox TxtCodiceAvvocato 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "XXX"
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2850
         TabIndex        =   2
         Top             =   225
         Width           =   285
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta :"
         Height          =   255
         Left            =   270
         TabIndex        =   32
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label LblDescrCodAvv 
         Caption         =   "Descrizione:"
         DataField       =   "NOME"
         Height          =   255
         Left            =   3330
         TabIndex        =   31
         Tag             =   "XXX"
         Top             =   250
         Width           =   6420
      End
      Begin VB.Label LblCodiceA 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1575
         TabIndex        =   30
         Top             =   270
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmDecretiIngiuntivi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim numOrdinamento As Integer
Dim codTribunale As String
Dim prezzoFormula As String
Private m_IsLoading As Boolean
Private moFrmRicerca As FrmRicerca
Public Azione As TipoAzione
Private sWhere As String
Private m_ID As Long

Implements IAnagraficForm
Implements IForm


Private Sub ChkAnnullo_Click()
    If ChkAnnullo.value = Checked Then
        LblAvvDecAnn.Visible = True
    Else
        LblAvvDecAnn.Visible = False
    End If
End Sub

Private Sub ChkEsenzione_Click()
        LblFormulaEsec.Visible = True
        ChkFormulaEsec.Visible = True
        fraEsenti.Visible = (ChkEsenzione.value = Unchecked)
         
    If ChkEsenzione.value = Checked Then
        txtSpese(2).value = 0
        TxtQta(2).value = 0
        txtSpese(3).value = 0
        TxtQta(3).value = 0
        
    End If
End Sub


Private Sub chkEvadi_Click()
If chkEvadi = 1 Then
 If Not m_IsLoading Then txtDataEvaso = Format(Date, "dd/mm/yyyy")
 LblDescrEvaso.Caption = "<< Decreto evaso"
 LblDescrEvaso.Visible = True
  txtSiglaCH.Tag = "necessario Sigla Chiusura"
Else
  txtDataEvaso = ""
  LblDescrEvaso.Visible = False
   txtSiglaCH.Tag = "Sigla Chiusura"
End If

End Sub

Private Sub ChkFormulaEsec_Click()
   If m_IsLoading Then Exit Sub
   
   fraFormula.Visible = (ChkFormulaEsec.value = Checked)
   InserisciPredefiniti
    If ChkFormulaEsec.value = Checked Then
        txtFormula.value = prezzoFormula
    Else
        txtFormula.value = 0
    End If
    
    CalcolaSaldo
End Sub

Private Sub SettaTribunale(cod As String)
Dim SQL As String
Dim rs As ADODB.Recordset

SQL = "SELECT CodiceTribunale " & _
       "FROM Anticipi " & _
       "WHERE CodiceAttivita='D' AND CodiceAlternativo='" & cod & "'"
Set rs = GetADORecordset("Anticipi", "CodiceTribunale", "CodiceAttivita='D' AND CodiceAlternativo='" & cod & "'", g_Settings.DBConnection)

If Not rs Is Nothing Then
    If Not rs.EOF Then
      If rs.RecordCount = 1 Then
        Dim tribID As Long
        tribID = rs(0).value
        SelectItemInTDBCombo cmbTribunale, tribID
        codTribunale = tribID
      End If
    End If
  Else
    SelectItemInTDBCombo cmbTribunale, -1
End If
 
End Sub

Private Sub cmbAutorita_SelChange(Cancel As Integer)
  If Not m_IsLoading Then
     Debug.Print "Cambio Autorità: " & cmbAutorita.Columns(1).value

   Select Case cmbAutorita.Columns(1).value
     Case "L", "GE"
        LblEsenzione.Visible = False
        ChkEsenzione.Visible = False
        ChkEsenzione.value = Checked
     Case "G", "T"
        LblEsenzione.Visible = False
        ChkEsenzione.Visible = False
        ChkEsenzione.value = Unchecked
     Case Else
        LblEsenzione.Visible = True
        ChkEsenzione.Visible = True
        ChkEsenzione.value = Unchecked
    End Select
    
    SettaTribunale cmbAutorita.Columns(1).value
    InserisciPredefiniti

   End If
    


End Sub



Private Sub cmbTribunale_SelChange(Cancel As Integer)
 If Not m_IsLoading Then
    InserisciPredefiniti
    ChkFormulaEsec_Click
 End If
End Sub



Private Sub CmdAnnulla_Click()
If CmdSalva.Enabled Then DeLockRecord m_ID, "DecretiIngiuntivi"
Unload Me
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
    cmbAutorita = ""
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

Private Sub CmdRicerca_Click()
    Set moFrmRicerca = New FrmRicerca
    
    Set moFrmRicerca.frmCaller = Me
    
    moFrmRicerca.tipo = "Ricerca"
    moFrmRicerca.Filtro = ""
    moFrmRicerca.Titolo = "Ricerca Decreti Ingiuntivi"
    moFrmRicerca.DefaultOrder = "Order By DataRegistrazione DESC, DecretiIngiuntivi.NumOrdinamento"
    moFrmRicerca.NCol = 7
    moFrmRicerca.PosizioneCodice = 8
    moFrmRicerca.Tabella = "DecretiIngiuntivi"
    moFrmRicerca.Query = "SELECT CheckVisual AS Ev, CODAVV AS [Codice], " & _
                "Format(Mid(DataRegistrazione,7,2) & '/' & Mid(DataRegistrazione,5,2) & '/' & Mid(DataRegistrazione,1,4),'dd/mm/yyyy') As [Data Registrazione], " & _
                " Ricorrente, Debitore,SIGLA as [Sigla Inserimento],SIGLACH as [Sigla chiusura], IDCod,NumeroDecreto, CodTribunaleApp, DataEvasionePratica,Annullo,DecretiIngiuntivi.NumOrdinamento FROM DecretiIngiuntivi "
'    If FindForm("frmRicerca") Then
'          Unload moFrmRicerca
'    End If
    Load moFrmRicerca

    'CmbTribunale.Enabled = False
End Sub

Private Sub CmdSalva_Click()
Dim Msg_Errore, contenuto1, contenuto2, contenuto3, contenuto4, contenuto5
Dim saved As Boolean
On Error GoTo ErroreSalvataggio
  If IsTableLocked("DecretiIngiuntivi") Then
       MsgBox "La tabelle decreti è bloccata da un altro utente. Riprovare...", vbInformation
  Else
        'LockTable ("DecretiIngiuntivi")
        SaveSetting "ATAP", "Config", "Sigla", txtSigla.Text
        saved = SalvaTutto(Me, "DecretiIngiuntivi", sWhere)
        
        If Not moFrmRicerca Is Nothing Then
            moFrmRicerca.AggiornaGriglia
        End If
        
        If saved Then DeLockRecord m_ID, "DecretiIngiuntivi"
        'DelockTable ("DecretiIngiuntivi")
        TxtCodiceAvvocato.SetFocus
  End If


Exit Sub

ErroreSalvataggio:

    If CmdSalva.Caption = "&Modifica" Then
        Msg_Errore = "Errore durante la modifica di un decreto "
    Else
        Msg_Errore = "Errore durante il salavataggio di un decreto "
    End If
    Msg_Errore = Msg_Errore & " - numero: " & err & " - riga: " & Erl & " - messaggio: " & Error(err)

    
    ErrLogFile "ErroriAtap.txt", Msg_Errore, contenuto1, contenuto2, contenuto3, contenuto4, contenuto5
    

End Sub

Private Sub Form_Load()
    m_IsLoading = True
    Me.Top = 0
   Azione = TipoAzione.Vuoto
   Call TipoMaschera(Me, Azione)
   txtDataReg.MaxDate = Now + 30
    PopolaTDBCombo cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale", "CodiceTribunale", , , "DescrizioneTribunale"
    PopolaTDBCombo cmbAutorita, "Autorita", "Codice as [Autorità]", "Codice"
    
    ChkFormulaEsec.value = Unchecked
    m_IsLoading = False
    
    
End Sub


Private Property Let IForm_IsLoading(RHS As Boolean)
 m_IsLoading = RHS
 Debug.Print "Is Loading: " & m_IsLoading
End Property
Private Property Get IForm_IsLoading() As Boolean
IForm_IsLoading = m_IsLoading
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


SQL = "SELECT DecretiIngiuntivi.CODAVV, " & _
      "( Mid(DataRegistrazione,7,2) & '/' & Mid(DataRegistrazione,5,2)& '/' & Mid(DataRegistrazione,1,4)) As DataRegistrazione, " & _
      "NumeroDecreto,NumeroIngiunzione,codAutorita, CodTribunaleApp, AnagraficaAvvocati.NOME, AnagraficaAvvocati.NumOrdinamento, " & _
      "ImpDepositoE, ImpCopieE,  QtaCopie, " & _
      "ImpFotocopieE, QtaFotocopie, " & _
      "ImpMarcheE,  QtaMarche, ImpFormulaE, " & _
      " ImpSpeseE, DesrSpese,  " & _
       " Debitore, Ricorrente,  " & _
       " Esenzione, FormulaEsec,QtaDirittiCancelleria,Commento,  " & _
       "DecretiIngiuntivi.NumOrdinamento,NumeroRuolo,NumeroDecreto," & _
      "ImpCompetenzeE, ImpSaldoE, " & _
      "( Mid(DataEvasionePratica,7,2) & '/' & Mid(DataEvasionePratica,5,2)& '/' & Mid(DataEvasionePratica,1,4)) As DataEvasionePratica, " & _
      "( Mid(DataDecreto,7,2) & '/' & Mid(DataDecreto,5,2)& '/' & Mid(DataDecreto,1,4)) As DataDecreto, Annullo,CheckVisual,SIGLA,SIGLACH, IDcod " & _
      "FROM (DecretiIngiuntivi INNER JOIN AnagraficaAvvocati ON DecretiIngiuntivi.CODAVV = AnagraficaAvvocati.CODAVV) INNER JOIN TribunaliAppartenenza ON DecretiIngiuntivi.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale " & _
      "WHERE " & sWhere
rs.Open SQL, g_Settings.DBConnection
m_ID = -1

If Not rs.EOF Then
   Azione = TipoAzione.Modifica
   
  
  
   Call TipoMaschera(Me, Azione)
   m_ID = rs("IDCod")
   If IsRecordLocked("IDCod=" & m_ID, "DecretiIngiuntivi") Then
      CmdSalva.Enabled = False
     Else
      CmdSalva.Enabled = True
      LockRecord m_ID, "DecretiIngiuntivi"
   End If
   
    Caricacampi Me, rs
   
 Else
    MsgBox "Il caricamento non è andato a buon fine:" & vbCrLf & "potrebbe non essere presente la Cassetta o il Tribunale corrispondente", vbCritical, "Attenzione"
End If
    


End Sub

Private Sub Timer1_Timer()
'CmdSalva.Enabled = Not IsRecordLocked("IDCod=" & m_ID, "DecretiIngiuntivi")
End Sub

Private Sub TxtCodiceAvvocato_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then CmdRicercaA_Click
End Sub

Private Sub InserisciPredefiniti()
 Dim SQL As String
 Dim rs As ADODB.Recordset
 Dim CodAutorita
 Debug.Print "Carica Predefiniti"
 
 
 codTribunale = cmbTribunale.Columns(1).value
 CodAutorita = cmbAutorita.Columns(1).value
 
 SQL = "SELECT * " & _
       "FROM Anticipi " & _
       "WHERE CodiceAttivita='D' AND CodiceAlternativo='" & CodAutorita & "' And CodiceTribunale='" & codTribunale & "'"
  Set rs = newAdoRs
  
  rs.Open SQL, g_Settings.DBConnection
  If Not rs.EOF Then
                    txtDeposito.value = rs!PrezDepositoEuro
                    txtCompetenze.value = rs!PrezCompetenzeEuro
                    prezzoFormula = rs!PrezFormulaEuro * Abs(CodAutorita = "L" Or (CodAutorita = "GE" And ChkEsenzione = Checked))
                    txtFormula.value = rs!PrezFormulaEuro * Abs(CodAutorita = "L" Or (CodAutorita = "GE" And ChkEsenzione = Checked))
                    txtSpese(1).value = rs!PrezFotocopieEuro
                    txtSpese(2).value = rs!PrezMarcheEuro * Abs(CodAutorita = "L" Or (CodAutorita = "GE" And ChkEsenzione = Checked))
                    txtSpese(3).value = rs!PrezCancelleriaEuro
  
      Else
                    txtDeposito.value = 0
                    txtCompetenze.value = 0
                    prezzoFormula = 0
                    txtFormula.value = 0
                    txtSpese(1).value = 0
                    txtSpese(2).value = 0
                    txtSpese(3).value = 0
  End If
  
  
 
  rs.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
'If FindForm("frmRicerca") Then
'    Unload FrmRicerca
'End If
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

Private Sub TxtFormula_Change()

    If Not m_IsLoading Then Call CalcolaSaldo
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub TxtQta_Change(Index As Integer)
   Call CalcolaSaldo
End Sub


Public Sub CalcolaSaldo()
On Error GoTo FINE
Dim saldo As Double
Dim spese As Double
Dim i As Integer
For i = 1 To 3
  spese = spese + txtSpese(i).value * TxtQta(i).value
Next
spese = spese + txtSpese(4).value

lblTotSpese(0).Caption = Format(spese, "##,##0.00")
lblTotSpese(1).Caption = Format(spese, "##,##0.00")

saldo = txtDeposito.value
saldo = saldo - txtFormula.value
saldo = saldo - txtCompetenze.value
saldo = saldo - spese

formattaSaldo LblValSaldo, saldo
Exit Sub
FINE:

End Sub





Private Sub txtSpese_Change(Index As Integer)
CalcolaSaldo
End Sub

Private Sub txtSpese_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Function IAnagraficForm_GetCodiceAvvocato() As String
  IAnagraficForm_GetCodiceAvvocato = TxtCodiceAvvocato.Text
End Function

Private Sub IAnagraficForm_RisultatoRicerca(sCodAvv As String, oAzione As TipoAzione)
Dim rs As ADODB.Recordset
Azione = TipoAzione.Nuovo
Set rs = GetADORecordset("AnagraficaAvvocati", "CodAvv,Nome,numOrdinamento", "CodAvv='" & sCodAvv & "'", g_Settings.DBConnection)
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
