VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaGiornalieraSfPig 
   Caption         =   "Stampa Giornaliera Sfratti Pignoramenti"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport CrptSfrattiPig 
      Left            =   1980
      Top             =   1665
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaGiornalieraSfPig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetActiveForm("StampaGiornalieraSfPig")
    gestioneReport
End Sub

Private Sub gestioneReport()
    setFileReport
    CrptSfrattiPig.WindowParentHandle = Me.hWnd
    Me.Move 0, 0, Atap.ScaleWidth, Atap.ScaleHeight
    
    CrptSfrattiPig.WindowState = crptMaximized
    CrptSfrattiPig.Destination = crptToWindow
    'CrptSfrattiPig.Action = 1
    CrptSfrattiPig.PrintReport
End Sub

Private Sub setFileReport()
    CrptSfrattiPig.DataFiles(0) = gDbName
    CrptSfrattiPig.ReportFileName = gPathReport & "\GiornalieraSfrattiPig.rpt"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call UnloadActiveForm("StampaGiornalieraSfPig")
End Sub



