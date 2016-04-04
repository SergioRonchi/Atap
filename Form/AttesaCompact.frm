VERSION 5.00
Begin VB.Form AttesaCompact 
   Caption         =   "Avviso"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   750
      Left            =   240
      Picture         =   "AttesaCompact.frx":0000
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Attendere compattazione dei database Access in corso ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "AttesaCompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
