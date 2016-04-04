VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaGiornalieraNotif 
   Caption         =   "Stampa Giornaliera Notifiche"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport CrptNotifiche 
      Left            =   1395
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaGiornalieraNotif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetActiveForm("StampaGiornalieraNotif")
    gestioneReport
End Sub

Private Sub gestioneReport()
    setFileReport
    CrptNotifiche.WindowParentHandle = Me.hWnd
    Me.Move 0, 0, Atap.ScaleWidth, Atap.ScaleHeight
    
    CrptNotifiche.WindowState = crptMaximized
    CrptNotifiche.Destination = crptToWindow
    'CrptNotifiche.Action = 1
    CrptNotifiche.PrintReport
End Sub

Private Sub setFileReport()
    CrptNotifiche.DataFiles(0) = gDbName
    CrptNotifiche.ReportFileName = gPathReport & "\GiornalieraNotifiche.rpt"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call UnloadActiveForm("StampaGiornalieraNotif")
End Sub


