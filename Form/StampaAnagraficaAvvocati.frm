VERSION 5.00
Begin VB.Form StampaAnagraficaAvvocati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Anagrafica Avvocati"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4800
      TabIndex        =   10
      Top             =   2160
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   1050
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6045
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Compatta Ordinamento Cognome Nome"
         Height          =   735
         Index           =   2
         Left            =   2385
         TabIndex        =   3
         Top             =   225
         Width           =   1455
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Compatta Ordinamento : Codice Cassetta"
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Dettagliata"
         Height          =   240
         Index           =   0
         Left            =   4500
         TabIndex        =   4
         Top             =   500
         Width           =   1410
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   1380
   End
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
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6045
      Begin VB.TextBox txtCodice 
         DataField       =   "CODAVV"
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txtDescrizione 
         DataField       =   "NOME"
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   1
         Top             =   600
         Width           =   3870
      End
      Begin VB.Label LblRicCodAvvInt 
         Caption         =   "Codice Cassetta :"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label LblRicNome 
         Caption         =   "Descrizione :"
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1365
      End
   End
End
Attribute VB_Name = "StampaAnagraficaAvvocati"
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
Dim nomeReport As String
    createSelectionFormula
 
 If OptTipoStampa(0).value = True Then nomeReport = "AnagraficaDettagliata.rpt"
    
 If OptTipoStampa(1).value = True Then nomeReport = "AnagraficaCondensata.rpt"
    
 If OptTipoStampa(2).value = True Then nomeReport = "AnagraficaC_Nome.rpt"
    
    Call Stampa.gestioneReport("", qrySQL, 0, crptToWindow, nomeReport, 1)
End Sub




Private Sub createSelectionFormula()
    qrySQL = "{AnagraficaAvvocati.CODAVV} LIKE '" & TxtCodice.Text & "*'"
    qrySQL = qrySQL & " AND {AnagraficaAvvocati.NOME} LIKE '" & TxtDescrizione.Text & "*'"
    qrySQL = qrySQL & " AND {AnagraficaAvvocati.STAT} <> 'A'"
End Sub


