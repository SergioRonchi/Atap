VERSION 5.00
Begin VB.Form frmback 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "-"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   930
   End
   Begin VB.Image pctLogo 
      Height          =   1815
      Left            =   1320
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "frmback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim D As String
lblNome = " A.T.A.P. Service versione " & app.Major & "." & app.Minor & " "
D = Dir(app.Path & "\logo.*")
If D <> "" Then Set Me.pctLogo.Picture = LoadPicture(app.Path & "\" & D)

End Sub

Private Sub Form_Resize()

If Atap.ScaleHeight - 2000 > 0 And Atap.ScaleWidth - 2000 > 0 Then Me.Move 1000, 1000, Atap.ScaleWidth - 2000, Atap.ScaleHeight - 2000
pctLogo.Move (Me.Width - pctLogo.Width) / 2, (Me.Height - pctLogo.Height) / 2
lblNome.Move (Me.Width - lblNome.Width) / 2, 500

End Sub
