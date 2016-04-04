VERSION 5.00
Begin VB.Form frmEditProv 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Modifica Provincia"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   400
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Salva"
      Height          =   400
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtCode 
      Height          =   375
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Codice"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_isOk As Boolean
Private m_initCode As String
Private m_Code As String
Private m_Descrizione As String
Private Sub CmdAnnulla_Click()
Unload Me
End Sub
Private Sub cmdButton_Click(Index As Integer)
 m_isOk = False
 If Trim(txtCode.Text) <> "" And Trim(txtNome.Text) <> "" Then
   If Exist(txtCode.Text) And m_initCode <> txtCode.Text Then
     MsgBox "Esiste già la provincia con codice " + txtCode.Text, vbOKOnly + vbInformation
    Else
     m_isOk = True
     m_Code = txtCode.Text
     m_Descrizione = txtNome.Text
     Unload Me
   End If
 End If
End Sub
Private Function Exist(code As String) As Boolean
 Exist = Not GetADORecordset("Provincie", "CodiceProvincia", "CodiceProvincia='" & code & "'", g_Settings.DBConnection) Is Nothing

End Function
Public Sub LoadData(cod As String, name As String)
txtCode.Text = cod
txtNome.Text = name
m_initCode = cod
End Sub
Public Property Get IsOk() As Boolean
IsOk = m_isOk
End Property
Public Property Get Codice() As String
Codice = m_Code
End Property
Public Property Get Descrizione() As String
Descrizione = m_Descrizione
End Property

