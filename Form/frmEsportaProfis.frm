VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2805E253-0A85-11D5-912D-9A7F711ED605}#1.1#0"; "MsgBoxEx.ocx"
Begin VB.Form frmEsportaProfis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esportazione"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Intervallo numeri di fattura"
      Height          =   855
      Left            =   480
      TabIndex        =   17
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox chkEnable 
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin TDBNumber6Ctl.TDBNumber tdbDa 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   564
         Calculator      =   "frmEsportaProfis.frx":0000
         Caption         =   "frmEsportaProfis.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEsportaProfis.frx":008C
         Keys            =   "frmEsportaProfis.frx":00AA
         Spin            =   "frmEsportaProfis.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "########0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbA 
         Height          =   315
         Left            =   4080
         TabIndex        =   21
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   564
         Calculator      =   "frmEsportaProfis.frx":011C
         Caption         =   "frmEsportaProfis.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEsportaProfis.frx":01A8
         Keys            =   "frmEsportaProfis.frx":01C6
         Spin            =   "frmEsportaProfis.frx":0210
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "########0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbAnno 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   564
         Calculator      =   "frmEsportaProfis.frx":0238
         Caption         =   "frmEsportaProfis.frx":0258
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEsportaProfis.frx":02C4
         Keys            =   "frmEsportaProfis.frx":02E2
         Spin            =   "frmEsportaProfis.frx":032C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
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
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label lblM 
         Caption         =   "Anno"
         DataField       =   "c"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Tag             =   "dal"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblM 
         Caption         =   "Da numero"
         DataField       =   "c"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   19
         Tag             =   "dal"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblM 
         Caption         =   "A numero"
         DataField       =   "c"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   18
         Tag             =   "al"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.OptionButton optMode 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optMode 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.Frame fraData 
      Caption         =   "Intervallo date"
      Height          =   855
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   6855
      Begin VB.ComboBox cmbDayNav 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   10
         Top             =   480
         Width           =   2160
      End
      Begin TDBDate6Ctl.TDBDate mskDal 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2734
         _ExtentY        =   450
         Calendar        =   "frmEsportaProfis.frx":0354
         Caption         =   "frmEsportaProfis.frx":046C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEsportaProfis.frx":04D8
         Keys            =   "frmEsportaProfis.frx":04F6
         Spin            =   "frmEsportaProfis.frx":0554
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   1
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "21/01/2003"
         ValidateMode    =   0
         ValueVT         =   2119368711
         Value           =   37642
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate mskAl 
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2734
         _ExtentY        =   450
         Calendar        =   "frmEsportaProfis.frx":057C
         Caption         =   "frmEsportaProfis.frx":0694
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEsportaProfis.frx":0700
         Keys            =   "frmEsportaProfis.frx":071E
         Spin            =   "frmEsportaProfis.frx":077C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   1
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "21/01/2003"
         ValidateMode    =   0
         ValueVT         =   2119368711
         Value           =   37642
         CenturyMode     =   0
      End
      Begin VB.Label lblM 
         Caption         =   "Al:"
         DataField       =   "c"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Tag             =   "al"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblM 
         Caption         =   "Dal:"
         DataField       =   "c"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Tag             =   "dal"
         Top             =   240
         Width           =   495
      End
   End
   Begin TDBNumber6Ctl.TDBNumber tdbNumFat 
      Height          =   320
      Left            =   3360
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   564
      Calculator      =   "frmEsportaProfis.frx":07A4
      Caption         =   "frmEsportaProfis.frx":07C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEsportaProfis.frx":0830
      Keys            =   "frmEsportaProfis.frx":084E
      Spin            =   "frmEsportaProfis.frx":0898
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100000
      MinValue        =   20
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1997471749
      Value           =   20
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00C0E0FF&
      DataField       =   "Giornate"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   2670
      Width           =   6840
   End
   Begin VB.CommandButton cmdSettings 
      Height          =   330
      Left            =   7065
      Picture         =   "frmEsportaProfis.frx":08C0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   390
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MsgBoxEx.IeButton cmdButton 
      Height          =   405
      Index           =   14
      Left            =   6240
      TabIndex        =   0
      Tag             =   "Esci"
      Top             =   4065
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   "Esci"
      ShowFocus       =   0   'False
   End
   Begin MsgBoxEx.IeButton cmdButton 
      Height          =   405
      Index           =   18
      Left            =   5115
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4065
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   "OK"
   End
   Begin VB.Label Label2 
      Caption         =   "Numero massimo di fatture per file"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Tag             =   "EXPORT.Percorso"
      Top             =   2400
      Width           =   3240
   End
End
Attribute VB_Name = "frmEsportaProfis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim key As String
Private WithEvents m_exporter As CExporter
Attribute m_exporter.VB_VarHelpID = -1



Private Sub chkEnable_Click()
 tdbA.Enabled = chkEnable.value = 1 And optMode(1).value
 If chkEnable.value = 0 Then
   tdbA.value = tdbA.MaxValue
  Else
   tdbA.value = tdbDa.value + 10000
 End If
End Sub

Private Sub cmbDayNav_Click()
  SettaDate mskDal, mskAl, cmbDayNav.ListIndex
 
End Sub

Private Sub cmdButton_Click(index As Integer)
On Error GoTo fine
Select Case index
  Case 18 'OK
   If txtPath.Text <> "" Then
     Dim D As String
     D = Format(Now, "YYYYMMDDHHmm")
     
     If optMode(1).value Then
       If tdbA.value < tdbDa.value Then
         MsgBox "L'intervallo del numero di fatture non è corretto", vbOKOnly + vbExclamation, "Atap"
         Exit Sub
       End If
     Else
     End If
     If Esporta(D) Then
      Dim zipUtils As New FileBackuoHelper
        Dim exportFile As String
        exportFile = "Atap_Export_" & D & ".zip"
        zipUtils.zipFile txtPath.Text, "Atap_Export_" & D & "_" & g_Settings.Sede, txtPath.Text, "*" & D & "*.csv", "Esportato file " & exportFile & " in "
        
        SafeKill txtPath.Text & "\*" & D & "*.csv"
        
        exportFile = "Atap_Export_XML" & D & ".zip"
        zipUtils.zipFile txtPath.Text, "Atap_Export_XML" & D & "_" & g_Settings.Sede, txtPath.Text, "*" & D & "*.xml", "Esportato file " & exportFile & " in "
        
        SafeKill txtPath.Text & "\*" & D & "*.xml"
        
        
        Shell "explorer.exe /e, " & txtPath.Text, vbNormalFocus
        Unload Me
     End If
     
     
   End If
  Case 14 'esci
   Unload Me
  
End Select
Exit Sub
fine:
  MsgBox err.Description, vbOKOnly + vbCritical
End Sub
Private Function Esporta(datePostFix As String) As Boolean
On Error GoTo fine
Dim nextMin As Long
Dim nextMax As Long

Dim oMinmax As MinMax
Dim esportaDatev As Boolean
Dim esportaXML As Boolean

ProgressBar1.Visible = True
ProgressBar2.Visible = True
Set oMinmax = New MinMax

esportaDatev = m_exporter.Esporta(oMinmax, txtPath.Text, optMode(1).value, _
                      tdbAnno.value, tdbDa.value, IIf(chkEnable.value = 1, tdbA.value, 1000000), mskDal.value, mskAl.value, datePostFix, tdbNumFat.value)
esportaXML = m_exporter.esportaXML(txtPath.Text, optMode(1).value, _
                      tdbAnno.value, tdbDa.value, IIf(chkEnable.value = 1, tdbA.value, 1000000), mskDal.value, mskAl.value, datePostFix)

If esportaDatev And esportaXML Then
  Esporta = True
 Else
  Esporta = False
End If

If optMode(1).value Then
  nextMin = oMinmax.IntMax + 1
  nextMax = IIf(chkEnable.value = 1, nextMin + (tdbA.value - tdbDa.value), nextMin + 10000)
  
  tdbDa.value = nextMin
  tdbA.value = nextMax
  
  SaveSetting "ATAP", "Export", key & "_MIN", nextMin
  SaveSetting "ATAP", "Export", key & "_MAX", nextMax
End If


Exit Function
fine:
Esporta = False
End Function



Private Sub cmdSettings_Click()
Dim initFolder As String
initFolder = txtPath.Text
 pGetFolder txtPath
 g_Settings.ExportPath = txtPath.Text
End Sub

Private Sub Form_Load()

Set m_exporter = New CExporter
CaricaDayNav cmbDayNav
cmbDayNav.ListIndex = 3

txtPath.Text = g_Settings.ExportPath
optMode_Click (1)
key = year(Now)

tdbAnno.value = key

tdbNumFat.value = GetSetting("ATAP", "Export", "NUMFAT", "100")
tdbDa.value = GetSetting("ATAP", "Export", key & "_MIN", "1")

tdbA.value = GetSetting("ATAP", "Export", key & "_MAXX", "100000")

End Sub
Private Sub pGetFolder(txt As TextBox)
 Dim DirSelected As String

   DirSelected = BrowseFolder(Me.hwnd, "Seleziona Cartella")
   If DirSelected <> "" Then
     txt.Text = DirSelected
   End If
   
End Sub


Private Sub txtBack_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "ATAP", "Export", "NUMFAT", tdbNumFat.value

End Sub

Private Sub m_exporter_OnDoubleProgress(v1 As Long, v2 As Long)
 ProgressBar1.value = v1
 ProgressBar2.value = v2
End Sub

Private Sub optMode_Click(index As Integer)

   mskAl.Enabled = optMode(0).value
   mskDal.Enabled = optMode(0).value
   cmbDayNav.Enabled = optMode(0).value
   
   tdbAnno.Enabled = optMode(1).value
   tdbDa.Enabled = optMode(1).value
   tdbA.Enabled = optMode(1).value
   chkEnable.Enabled = optMode(1).value
End Sub
