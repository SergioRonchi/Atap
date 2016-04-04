VERSION 5.00
Begin VB.Form StampaAnticipi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Anticipi"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FmRicerca 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1050
      Left            =   50
      TabIndex        =   2
      Top             =   0
      Width           =   6045
      Begin VB.TextBox txtDescrizione 
         DataField       =   "NOME"
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   3870
      End
      Begin VB.TextBox txtCodice 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label LblRicNome 
         Caption         =   "Descrizione :"
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label LblRicCodAvvInt 
         Caption         =   "Codice Anticipo :"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1380
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   1380
   End
End
Attribute VB_Name = "StampaAnticipi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qrySQL As String

Private Sub CmdAnnulla_Click()
Unload Me
End Sub

Private Sub CmdOK_Click()
    createSelectionFormula
    Call Stampa.gestioneReport("", qrySQL, 0, crptToWindow, "AnticipiEuro.rpt", 1)
End Sub

Private Sub createSelectionFormula()
    qrySQL = "{Anticipi.CodiceAnticipi} LIKE """ & TxtCodice.Text & "*"""
    qrySQL = qrySQL & " AND {Anticipi.Descrizione} LIKE """ & TxtDescrizione.Text & "*"""
End Sub

