VERSION 5.00
Begin VB.Form frmApriStorico 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Gestione Estratti Conto"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optUNEP 
      Caption         =   "Estratti conto UNEP"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton optAdempi 
      Caption         =   "Estratti Conto Adempimenti"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.OptionButton optEC 
      Caption         =   "Estratti Conto"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame fraInfo 
      Caption         =   "info"
      Height          =   2175
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Label3 
         Caption         =   "del XX/XX/XXXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "del XX/XX/XXXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Estratto Conto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdApri 
      Caption         =   "Apri"
      Height          =   500
      Left            =   10440
      TabIndex        =   2
      Top             =   5880
      Width           =   1140
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   11760
      TabIndex        =   1
      Top             =   5880
      Width           =   1140
   End
   Begin VB.FileListBox File1 
      Height          =   5745
      Left            =   120
      Pattern         =   "*.mdb"
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "frmApriStorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Path As String
Public aperto As Boolean
Public codice As String

Private Sub CmdAnnulla_Click()
Unload Me
End Sub

Private Sub cmdApri_Click()
On Error GoTo FINE
If File1.ListIndex >= 0 Then
     Dim dbUpdater As CDBUpdater
     Set dbUpdater = New CDBUpdater
 
 g_Settings.ConnettiDB (g_Settings.dbPath & "Storico\" & Path & "\" & File1.List(File1.ListIndex))
 
 dbUpdater.UpdateDatabase g_Settings.DBConnection
 aperto = True
 MsgBox "Database Storico Aperto Correttamente", vbInformation
 
 Unload Me
 Else
  MsgBox "Nessun Archivio selezionato", vbCritical
End If
Exit Sub
FINE:
 MsgBox err.Description, vbCritical
End Sub

Private Sub File1_Click()
 Dim f As String
 Dim n As String
 Dim pos As Integer
 If File1.ListIndex >= 0 Then
    f = File1.List(File1.ListIndex)
    If Path = "EstrattiConto" Then
       Label1 = "Estratto Conto"
       
         f = Mid(f, 4, 8)
       
       
       Label2 = "del " & RitornaData(f)
       Label3 = ""
     ElseIf Path = "EstrattiConto\Adempimenti" Then
       Label1 = "Estratto Conto Adempimenti"
       f = Mid(f, 11, 8)
     ElseIf Path = "EstrattiConto\UNEP" Then
       Label1 = "Estratto Conto UNEP"
       f = Mid(f, 8, 8)
       Label2 = "del " & RitornaData(f)
       Label3 = ""
     Else
     
      Label1 = "Liquidazione"
      pos = InStr(1, f, "_")
      n = Mid(f, 4, pos - 4)
      codice = n
      Label2 = "Cassetta: " & n
      f = Mid(f, pos + 1, 8)
      Label3 = " del " & RitornaData(f)

    End If
    
    fraInfo.Visible = True
  Else
    fraInfo.Visible = False
 End If
 
End Sub

Private Sub Form_Load()
aperto = False
File1.Path = g_Settings.dbPath & "Storico\" & Path & "\"
If Path <> "EstrattiConto" Then
  optEC.Visible = False
  optAdempi.Visible = False
End If
End Sub

Private Sub optAdempi_Click()
 Path = "EstrattiConto\Adempimenti"
 File1.Path = g_Settings.dbPath & "Storico\" & Path & "\"
End Sub

Private Sub optEC_Click()
 Path = "EstrattiConto"
 File1.Path = g_Settings.dbPath & "Storico\" & Path & "\"
End Sub

Private Sub optUNEP_Click()
 Path = "EstrattiConto\UNEP"
 File1.Path = g_Settings.dbPath & "Storico\" & Path & "\"
End Sub
