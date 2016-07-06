VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2805E253-0A85-11D5-912D-9A7F711ED605}#1.1#0"; "MsgBoxEx.ocx"
Begin VB.Form frmEsportaProfis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esportazione"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      BackColor       =   &H00C0E0FF&
      DataField       =   "Giornate"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   870
      Width           =   6840
   End
   Begin VB.CommandButton cmdSettings 
      Height          =   330
      Left            =   6945
      Picture         =   "frmEsportaProfis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   390
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
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
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
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
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   2160
   End
   Begin TDBDate6Ctl.TDBDate mskDal 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2734
      _ExtentY        =   450
      Calendar        =   "frmEsportaProfis.frx":02B2
      Caption         =   "frmEsportaProfis.frx":03CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEsportaProfis.frx":0436
      Keys            =   "frmEsportaProfis.frx":0454
      Spin            =   "frmEsportaProfis.frx":04B2
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
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2734
      _ExtentY        =   450
      Calendar        =   "frmEsportaProfis.frx":04DA
      Caption         =   "frmEsportaProfis.frx":05F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEsportaProfis.frx":065E
      Keys            =   "frmEsportaProfis.frx":067C
      Spin            =   "frmEsportaProfis.frx":06DA
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
   Begin MsgBoxEx.IeButton cmdButton 
      Height          =   405
      Index           =   14
      Left            =   6120
      TabIndex        =   4
      Tag             =   "Esci"
      Top             =   2265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   "Esci"
      ShowFocus       =   0   'False
   End
   Begin MsgBoxEx.IeButton cmdButton 
      Height          =   405
      Index           =   18
      Left            =   4995
      TabIndex        =   5
      Tag             =   "OK"
      Top             =   2265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   "OK"
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Tag             =   "EXPORT.Percorso"
      Top             =   600
      Width           =   3240
   End
   Begin VB.Label lblM 
      Caption         =   "Dal:"
      DataField       =   "c"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Tag             =   "dal"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblM 
      Caption         =   "Al:"
      DataField       =   "c"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   2
      Tag             =   "al"
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmEsportaProfis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_exporter As CExporter
Attribute m_exporter.VB_VarHelpID = -1



Private Sub cmbDayNav_Click()
  SettaDate mskDal, mskAl, cmbDayNav.ListIndex
 
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo fine
Select Case Index
  Case 18 'OK
   If txtPath.Text <> "" Then
     If Esporta() Then
       
       MsgBox "File esportati", vbOKOnly + vbInformation
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
Private Function Esporta() As Boolean
On Error GoTo fine


ProgressBar1.Visible = True
ProgressBar2.Visible = True

If m_exporter.Esporta(txtPath.Text, mskDal.value, mskAl.value) Then
  Esporta = True
 Else
  Esporta = False
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
mskDal.value = "2000-01-01"
txtPath.Text = g_Settings.ExportPath


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

Private Sub m_exporter_OnDoubleProgress(v1 As Long, v2 As Long)
 ProgressBar1.value = v1
 ProgressBar2.value = v2
End Sub
