VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPathDB 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imposta il percorso del database"
   ClientHeight    =   2805
   ClientLeft      =   3030
   ClientTop       =   2835
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fratab 
      Caption         =   " Cerca il database:"
      ForeColor       =   &H00800000&
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7695
      Begin MSComDlg.CommonDialog CMDSfoglia1 
         Left            =   240
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSfoglia 
         Caption         =   "Sfoglia..."
         Height          =   350
         Left            =   6360
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "frmPathDB.frx":0000
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Indicare il percorso di ubicazione del database Atap.mdb"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Cerca in:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAnnulla 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   350
      Left            =   6480
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmPathDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Dbname As String
Dim mExitMode As enumExitMode
Dim mDaiMessaggio As Boolean
Private m_dbPath As String
Public Property Get DBName() As String
  DBName = m_Dbname
End Property
Private Sub CmdAnnulla_Click()
    mExitMode = exitCANCEL
    Unload Me
    MsgBox "Il programma non riesce a trovare o ad aprire il database Atap.mdb corretto." & vbCrLf & _
           "Verificare il percorso del database." & vbCrLf & _
           "Se il problema persiste il database potrebbe essere rovinato. Conttattare quindi il fornitore del software", vbInformation, "Attenzione"
End Sub

Private Sub CmdOK_Click()
'se non modifico il path del DB, allora mantengo quello impostato
    If txtPath.Text <> "" Then
        m_Dbname = txtPath.Text
        
        
End If
    
'non posso chiamare la routine del Registro perchè non essendosi ancora caricata, mi dà errore
   mExitMode = exitOk
   
   Unload Me
   
    
End Sub

Private Sub cmdSfoglia_Click()
'sfrutto la commond dialog di W95 per trovare il mio DB
    CMDSfoglia1.InitDir = m_dbPath
    CMDSfoglia1.ShowOpen
    If CMDSfoglia1.fileName <> "" Then
       txtPath = CMDSfoglia1.fileName
    End If

End Sub


Public Sub Initialize(dbFile As String, dbPath As String)
txtPath = dbFile
m_dbPath = dbPath
End Sub

Public Property Get ExitMode() As enumExitMode
'il metodo Property Get restituisce un valore (Es. ExitMode)
    ExitMode = mExitMode
End Property

Public Property Let ExitMode(ByVal vNewValue As enumExitMode)
'il metodo Property Let imposta un valore (Es. mExitMode)
    mExitMode = vNewValue
End Property

