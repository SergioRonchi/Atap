VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StampaEstrattoDefinitiva 
   Caption         =   "Gestione stampa Estratto Conto Definitiva"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   5280
   Begin Crystal.CrystalReport CRptEstrattoDefinitivo 
      Left            =   1800
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "StampaEstrattoDefinitiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    
    StampaEstrattoConto.Enabled = False
    gestioneReport
    Screen.MousePointer = vbDefault
End Sub

Private Sub gestioneReport()
    setFileReport
    CRptEstrattoDefinitivo.WindowParentHandle = Me.hWnd
    Me.Move 0, 0, Atap.ScaleWidth, Atap.ScaleHeight
    
    CRptEstrattoDefinitivo.WindowState = crptMaximized
    CRptEstrattoDefinitivo.Destination = crptToWindow
    CRptEstrattoDefinitivo.PrintReport
End Sub

Private Sub setFileReport()
    CRptEstrattoDefinitivo.DataFiles(0) = g_Settings.DBFile
    CRptEstrattoDefinitivo.ReportFileName = g_Settings.ReportPath & "\EstrattoConto.rpt"
    CRptEstrattoDefinitivo.Formulas(0) = "Tipo='ESTRATTO'"
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Prosegui
    StampaEstrattoConto.Enabled = True
End Sub

Public Sub Prosegui()

Dim MSG_Avviso, Response As Variant
    
    MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
    MSG_Avviso = MSG_Avviso & "Si vuole Procedere con la creazione della stampa Richiesta assegni circolari?" & Chr(10)
    MSG_Avviso = MSG_Avviso & "(Obbligatorio per rendere definitivo l'estratto conto)"
    Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
    If Response = vbYes Then    ' User chose Yes.
        Screen.MousePointer = vbHourglass
        StampaEstrattoConto.CreazioneStampaAssegniCircolari
        Screen.MousePointer = vbDefault
        CRptEstrattoDefinitivo.DataFiles(0) = g_Settings.DBFile
        CRptEstrattoDefinitivo.Formulas(0) = ""
        CRptEstrattoDefinitivo.ReportFileName = g_Settings.ReportPath & "\AssegniCircolari.rpt"
        CRptEstrattoDefinitivo.Destination = crptToPrinter
        CRptEstrattoDefinitivo.CopiesToPrinter = 3
        CRptEstrattoDefinitivo.PrintReport
        MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
        MSG_Avviso = MSG_Avviso & "Si vuole Procedere col trasferimento dei dati nel database storico?" & Chr(10)
        MSG_Avviso = MSG_Avviso & "(Obbligatorio per rendere definitivo l'estratto conto)"
        Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
        If Response = vbYes Then    ' User chose Yes.
            Screen.MousePointer = vbHourglass
'            StampaEstrattoConto.TrasferimentoDatiAlDbStorico
'            If StampaEstrattoConto.TrasferimentoOK = True Then
'                StampaEstrattoConto.UpdateDataUltimoEstConto
'                StampaEstrattoConto.EliminazioneDatiTrasferiti
'                StampaEstrattoConto.AggiornaTabellaSaldi
'                Screen.MousePointer = vbDefault
'                ImpostazioniFatturazione.Show
'            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End Sub
