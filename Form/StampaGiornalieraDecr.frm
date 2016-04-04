VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaGiornalieraDecr 
   Caption         =   "Stampa Giornaliera Decreti Ingiuntivi"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport CrptDecretiIng 
      Left            =   2115
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaGiornalieraDecr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetActiveForm("StampaGiornalieraDecr")
    gestioneReport
End Sub

Private Sub gestioneReport()
    setFileReport
    CrptDecretiIng.WindowParentHandle = Me.hWnd
    Me.Move 0, 0, Atap.ScaleWidth, Atap.ScaleHeight
    
    CrptDecretiIng.WindowState = crptMaximized
    CrptDecretiIng.Destination = crptToWindow
    'CrptDecretiIng.Action = 1
    CrptDecretiIng.PrintReport
End Sub

Private Sub setFileReport()
    CrptDecretiIng.DataFiles(0) = gDbName
    CrptDecretiIng.ReportFileName = gPathReport & "\GiornalieraDecreti.rpt"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call UnloadActiveForm("StampaGiornalieraDecr")
End Sub



