VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmConfigurazione 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurazione"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIBAN 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox txtBanca 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5655
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
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
      Top             =   240
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
      ValueVT         =   142409729
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
   Begin VB.Label Label5 
      Caption         =   "IBAN"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Banca"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
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
      Top             =   0
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
 Unload Me
End Sub

Private Sub Form_Load()
    tdbIVA.value = g_Settings.IVA
    tdbSoci.value = g_Settings.QuotaSoci
    numLimitesaldo.value = g_Settings.LimiteSaldo
    txtBanca.Text = g_Settings.Banca
    txtIBAN.Text = g_Settings.IBAN
End Sub
