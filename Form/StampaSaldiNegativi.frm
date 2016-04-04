VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form StampaSaldiNegativi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stampa Saldi Precedenti Negativi"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1440
      TabIndex        =   12
      Top             =   2565
      Width           =   1050
   End
   Begin VB.CommandButton CmdPulisci 
      Caption         =   "&Pulisci"
      Height          =   330
      Left            =   2655
      TabIndex        =   11
      Top             =   2565
      Width           =   1050
   End
   Begin VB.Frame FrmProvvisoria 
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4920
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   285
         Left            =   2925
         TabIndex        =   5
         Top             =   810
         Width           =   1860
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         Top             =   810
         Width           =   1050
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2475
         TabIndex        =   3
         Top             =   810
         Width           =   330
      End
      Begin VB.TextBox TxtRicDataFin 
         Height          =   285
         Left            =   3690
         TabIndex        =   2
         Top             =   315
         Width           =   1035
      End
      Begin VB.TextBox TxtRicDataIn 
         Height          =   285
         Left            =   1305
         TabIndex        =   1
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label LblDescrCodAvv 
         Height          =   690
         Left            =   1395
         TabIndex        =   10
         Top             =   1350
         Width           =   3345
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Avvocato:"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label LblDescr 
         Caption         =   "Descrizione:"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
   End
   Begin Crystal.CrystalReport CRptSospesi 
      Left            =   4545
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaSaldiNegativi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

