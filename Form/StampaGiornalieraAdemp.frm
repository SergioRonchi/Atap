VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaGiornalieraAdemp 
   Caption         =   "Stampa Giornaliera Adempimenti Cancelleria"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport CRptAdempimenti 
      Left            =   1980
      Top             =   1485
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaGiornalieraAdemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetActiveForm("StampaGiornalieraAdemp")
    gestioneReport
End Sub

Private Sub gestioneReport()
    setFileReport
    CRptAdempimenti.WindowParentHandle = Me.hWnd
    Me.Move 0, 0, Atap.ScaleWidth, Atap.ScaleHeight
    
    CRptAdempimenti.WindowState = crptMaximized
    CRptAdempimenti.Destination = crptToWindow
    'CRptAdempimenti.Action = 1
    CRptAdempimenti.PrintReport
End Sub

Private Sub setFileReport()
    CRptAdempimenti.DataFiles(0) = gDbName
    CRptAdempimenti.ReportFileName = gPathReport & "\GiornalieraAdempimenti.rpt"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call UnloadActiveForm("StampaGiornalieraAdemp")
End Sub


