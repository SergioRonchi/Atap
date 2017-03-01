VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmConfigurazione 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurazione"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Codici del Piano dei Conti"
      Height          =   2895
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   8175
      Begin VB.TextBox txtCodTestataIncasso 
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodComp 
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtCodVar 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtCodFix 
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtCodTestata 
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Testata Incasso"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Competenze"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Quote Variabili"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Quote Associative Fisse"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Testata fattura"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox txtBackup 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   6615
   End
   Begin VB.TextBox txtCodIva 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtIBAN 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   5655
   End
   Begin VB.TextBox txtBanca 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   5655
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   5160
      TabIndex        =   3
      Top             =   6600
      Width           =   1500
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   6840
      TabIndex        =   2
      Top             =   6600
      Width           =   1500
   End
   Begin TDBNumber6Ctl.TDBNumber tdbIVA 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      Calculator      =   "frmConfigurazione.frx":0000
      Caption         =   "frmConfigurazione.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmConfigurazione.frx":008C
      Keys            =   "frmConfigurazione.frx":00AA
      Spin            =   "frmConfigurazione.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0.0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0.0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   142409729
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbSoci 
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      Calculator      =   "frmConfigurazione.frx":011C
      Caption         =   "frmConfigurazione.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmConfigurazione.frx":01A8
      Keys            =   "frmConfigurazione.frx":01C6
      Spin            =   "frmConfigurazione.frx":0210
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0.00;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   1000
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   75366401
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber numLimitesaldo 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      Calculator      =   "frmConfigurazione.frx":0238
      Caption         =   "frmConfigurazione.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmConfigurazione.frx":02C4
      Keys            =   "frmConfigurazione.frx":02E2
      Spin            =   "frmConfigurazione.frx":032C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0.00;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   1000
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   142409729
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label7 
      Caption         =   "Path per il backup automatico"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Codice IVA"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "IBAN"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Banca"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Limite Saldo"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Contributo bimestrale soci"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Aliquota IVA:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmConfigurazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAnnulla_Click()
Unload Me
End Sub

Private Sub CmdOK_Click()
  g_Settings.IVA = tdbIVA.value
  g_Settings.QuotaSoci = tdbSoci.value
  g_Settings.LimiteSaldo = numLimitesaldo.value
  g_Settings.Banca = txtBanca.Text
  g_Settings.IBAN = txtIBAN.Text
  g_Settings.CodIVA = txtCodIva.Text
  g_Settings.CodCompetenze = txtCodComp.Text
  g_Settings.CodTestata = txtCodTestata.Text
  g_Settings.CodQuotaVariabile = txtCodVar.Text
  g_Settings.CodQuataFissa = txtCodFix.Text
  g_Settings.CodTestataIncasso = txtCodTestataIncasso.Text
 Unload Me
End Sub

Private Sub Form_Load()
    tdbIVA.value = g_Settings.IVA
    tdbSoci.value = g_Settings.QuotaSoci
    numLimitesaldo.value = g_Settings.LimiteSaldo
    txtBanca.Text = g_Settings.Banca
    txtIBAN.Text = g_Settings.IBAN
    txtCodIva = g_Settings.CodIVA
    txtBackup = g_Settings.AtapUserBackupFolder
    txtCodTestata = g_Settings.CodTestata
    txtCodTestataIncasso = g_Settings.CodTestataIncasso
    txtCodFix = g_Settings.CodQuataFissa
    txtCodVar = g_Settings.CodQuotaVariabile
    txtCodComp = g_Settings.CodCompetenze
End Sub

