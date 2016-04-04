VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPathReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imposta il percorso del database storico"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2370
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnulla 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   350
      Left            =   4410
      TabIndex        =   5
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Frame fratab 
      Caption         =   " Cerca il database storico:"
      ForeColor       =   &H00800000&
      Height          =   2175
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5655
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton cmdSfoglia 
         Caption         =   "Sfoglia..."
         Height          =   350
         Left            =   4200
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CMDSfoglia1 
         Left            =   240
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Immettere il percorso dove sono salvati i report oppure fare click sul pulsante sfoglia."
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "frmPathReport.frx":0000
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmPathReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mExitMode As enumExitMode
Dim mDaiMessaggio As Boolean

Private Sub CmdAnnulla_Click()
    mExitMode = exitCANCEL
    Unload Me
    MsgBox "Per poter esguire le stampe bisogna selezionare il percorso dei report corretto!", vbInformation, "Attenzione"
End Sub

Private Sub CmdOK_Click()
'se non modifico il path del DB, allora mantengo quello impostato
    If txtPath.Text <> "" Then
        gDbStorico = txtPath.Text
        'Salva o crea una voce per un'applicazione nel registro di configurazione
        SaveSetting cNomeInRegistry, "path", "storico", gDbStorico
        'cNomeInRegistry. Espressione stringa contenente il nome dell'applicazione o del progetto a cui si riferisce l'impostazione.
        'path. Espressione stringa contenente il nome della sezione nella quale viene salvata l'impostazione di chiave.
        'database. Espressione stringa contenente il nome dell'impostazione di chiave salvata.
        'gDbName. Espressione contenente il valore sul quale viene impostato l'argomento key.
    End If
    
'non posso chiamare la routine del Registro perchè non essendosi ancora caricata, mi dà errore
   mExitMode = exitOk
   
   Unload Me
   'Atap.Show
    
End Sub

Private Sub cmdSfoglia_Click()
'sfrutto la commond dialog di W95 per trovare il mio DB

    CMDSfoglia1.ShowOpen
    If CMDSfoglia1.FileName <> "" Then
      ' txtPath = CMDSfoglia1.FileName
      txtPath = CMDSfoglia1.FileName
    End If

End Sub


Private Sub Form_Load()
'Assegna alla textbox il path del db
    txtPath = gDbStorico
End Sub


Public Property Get ExitMode() As enumExitMode
'il metodo Property Get restituisce un valore (Es. ExitMode)
    ExitMode = mExitMode
End Property

Public Property Let ExitMode(ByVal vNewValue As enumExitMode)
'il metodo Property Let imposta un valore (Es. mExitMode)
    mExitMode = vNewValue
End Property


