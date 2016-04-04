VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaAssCircProv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Assegni Circolari Provvisoria"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   3480
      TabIndex        =   15
      Top             =   2880
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Width           =   1380
   End
   Begin VB.Frame FrmMetodoStampa 
      Caption         =   "Modalità Stampa"
      Height          =   645
      Left            =   50
      TabIndex        =   10
      Top             =   2160
      Width           =   4800
      Begin VB.OptionButton OptModSt 
         Caption         =   "Anteprima"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton OptModSt 
         Caption         =   "Diretta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame FrmProvvisoria 
      Height          =   2115
      Left            =   50
      TabIndex        =   5
      Top             =   0
      Width           =   4800
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   330
      End
      Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "necessario Data Registrazione"
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaAssCircProv.frx":0000
         Caption         =   "StampaAssCircProv.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaAssCircProv.frx":0184
         Keys            =   "StampaAssCircProv.frx":01A2
         Spin            =   "StampaAssCircProv.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   ""
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   ""
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   2.12482833205922E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TxtRicDataFin 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Tag             =   "necessario Data Registrazione"
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaAssCircProv.frx":0228
         Caption         =   "StampaAssCircProv.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaAssCircProv.frx":03AC
         Keys            =   "StampaAssCircProv.frx":03CA
         Spin            =   "StampaAssCircProv.frx":0428
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   ""
         HighlightText   =   2
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   ""
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   2.12482833205922E-314
         CenturyMode     =   0
      End
      Begin VB.Label LblDescrCodAvv 
         Caption         =   "TUTTE LE CASSETTE"
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   4545
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta:"
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
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdPulisci 
      Caption         =   "&Pulisci"
      Height          =   500
      Left            =   2040
      TabIndex        =   4
      Top             =   2880
      Width           =   1380
   End
End
Attribute VB_Name = "StampaAssCircProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_Avvocato As String
Private WithEvents moFilterManager As CFilterManager
Attribute moFilterManager.VB_VarHelpID = -1

Private Sub CmdAnnulla_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If
End Sub

Private Sub CmdOK_Click()


Dim MSG_Avviso, Response As Variant
    
    If IsPrtTableLocked("PrtAssegniCircolari") Or IsPrtTableLocked("PrtEstrattoConto") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If
    LockPrtTable ("PrtAssegniCircolari")
    LockPrtTable ("PrtEstrattoConto")
    
    Riempi_PRT_EstrattoConto
    
   
    If GetADORecordset("PrtEstrattoConto", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
            MsgBox "Nessun dato evaso! Impossibile creare la stampa Elenco Assegni Circolari Provvisoria.", vbInformation, "Attenzione"
        GoTo Sblocca
    End If
  
    
    CreazioneStampaAssegniCircolari
    
    
    Call Stampa.gestioneReport("PrtAssegniCircolari", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "AssegniCircolari.rpt", 1)
    
    
Sblocca:
    CloseProgress
    DelockPrtTable ("PrtAssegniCircolari")
    DelockPrtTable ("PrtEstrattoConto")
    
     
    
End Sub

Private Sub CmdPulisci_Click()
    TxtCodiceAvvocato.Text = ""
    LblDescrCodAvv.Caption = ""
    TxtRicDataIn.value = #1/1/1999#
    TxtRicDataFin = ""
End Sub




Private Sub Form_Load()
    Set moFilterManager = New CFilterManager
    moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv
   
   CmdPulisci_Click
   m_Avvocato = "ALL"
End Sub

Private Sub moFilterManager_Validate(IsValid As Boolean)
   CmdOK.Enabled = IsValid
End Sub


Public Sub Riempi_PRT_EstrattoConto()
Dim qrySQL As String
Dim qryApp As String
Dim qryDelete As String
Dim qry1, qry2, qry3 As String
Dim NumErrori As Integer
    

On Error GoTo Riempi_PRT_EstrattoConto

  
    If IsDate(TxtRicDataIn) Then qry1 = " AND ( DataEvasionePratica >= '" & Format(TxtRicDataIn.Text, "yyyymmdd") & "')"
    If IsDate(TxtRicDataFin) Then qry2 = " AND ( DataEvasionePratica <= '" & Format(TxtRicDataFin, "yyyymmdd") & "')"
        
    If TxtCodiceAvvocato.Text <> "" Then
        qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & TxtCodiceAvvocato.Text & "')"
    End If
    
    qryApp = qry1 & qry2 & qry3
    
    OpenProgress ("Attendere... Preparazione Stampa!")
    UpdateProgress 0, "Adempimenti"
    'Reset PrtEstrattoConto
    g_Settings.DBConnection.Execute "DELETE * FROM PrtEstrattoConto;"
    
    'Inizio Adempimenti
    qrySQL = getQryAdempimenti(False, "Futuro", TxtRicDataIn.Text, TxtRicDataIn.Text, "S") & qryApp & " ORDER BY ADEMPI.NumOrdinamento"
    update_EstConto_Adempimenti "PrtEstrattoConto", qrySQL ', NumEstConto
    UpdateProgress 25, "Sfratti"
    'Fine Adempimenti

    'Inizio Sfratti
    qrySQL = getQrySfratti(False, "Futuro", TxtRicDataIn.Text, TxtRicDataIn.Text, "S") & qryApp & "  ORDER BY SFRATTI.NumOrdinamento"
    update_EstConto_Sfratti "PrtEstrattoConto", qrySQL
    UpdateProgress 50, "Notifiche"
    ' Fine Sfratti
    
    'Inizio Notifiche
    qrySQL = getQryNotifiche(False, "Futuro", TxtRicDataIn.Text, TxtRicDataIn.Text, "S") & qryApp & "  ORDER BY Notifiche.NumOrdinamento"
    update_EstConto_Notifiche "PrtEstrattoConto", qrySQL
    UpdateProgress 75, "Decreti"
    'Fine Notifiche

    'Inizio Decreti
    qrySQL = getQryDecreti(False, "Futuro", TxtRicDataIn.Text, TxtRicDataIn.Text, "S") & qryApp & " ORDER BY DecretiIngiuntivi.NumOrdinamento"
    update_EstConto_Decreti "PrtEstrattoConto", qrySQL
    UpdateProgress 100, "Stampa in corso..."
    'Fine Decreti
   

Exit Sub

Riempi_PRT_EstrattoConto:
    
        MsgBox "Attenzione errore in stampa Estratto Conto!" & Chr(10) & err & " - " & Error(err), vbCritical, "Attenzione"
 

End Sub

Public Sub CreazioneStampaAssegniCircolari()
On Error GoTo FINE
Dim SQL As String

SQL = "INSERT INTO PrtAssegniCircolari ( CODAVV, Nome, DEPOSITO, SPESE1, SPESE2, SPESE3, SPESE4, SPESE5, SPESE6, COMPETENZE, SALDO, SALDO_PRECEDENTE, DATA_INIZIO, DATA_FINE, NumOrdinamento,DESCR_ATTIVITA,Valuta ) " & _
      "SELECT PrtEstrattoConto.CODAVV, PrtEstrattoConto.Nome, Sum(PrtEstrattoConto.DEPOSITO) AS SommaDiDEPOSITO, Sum(PrtEstrattoConto.SPESE1) AS SommaDiSPESE1," & _
      "Sum(PrtEstrattoConto.SPESE2) AS SommaDiSPESE2, Sum(PrtEstrattoConto.SPESE3) AS SommaDiSPESE3, Sum(PrtEstrattoConto.SPESE4) AS SommaDiSPESE4," & _
      "Sum(PrtEstrattoConto.SPESE5) AS SommaDiSPESE5, Sum(PrtEstrattoConto.SPESE6) AS SommaDiSPESE6, Sum(PrtEstrattoConto.COMPETENZE) AS SommaDiCOMPETENZE," & _
      "Sum(PrtEstrattoConto.SALDO) AS actSaldo, First(PrtEstrattoConto.SALDO_PRECEDENTE) AS prevSaldo," & _
      "First(PrtEstrattoConto.DATA_INIZIO) AS PrimoDiDATA_INIZIO, First(PrtEstrattoConto.DATA_FINE) AS PrimoDiDATA_FINE," & _
      "First(AnagraficaAvvocati.NumOrdinamento) AS PrimoDiNumOrdinamento,' ','E' " & _
      "FROM PrtEstrattoConto INNER JOIN AnagraficaAvvocati ON PrtEstrattoConto.CODAVV = AnagraficaAvvocati.CODAVV " & _
      "GROUP BY PrtEstrattoConto.CODAVV, PrtEstrattoConto.Nome " & _
      "Having (((Sum([PrtEstrattoConto].[saldo]) + First([PrtEstrattoConto].[SALDO_PRECEDENTE])) >= " & Str(g_Settings.LimiteSaldo) & ")) " & _
      "ORDER BY First(AnagraficaAvvocati.NumOrdinamento);"

'Reset PrtEstrattoConto
g_Settings.DBConnection.BeginTrans
g_Settings.DBConnection.Execute "DELETE  * FROM PrtAssegniCircolari;"
g_Settings.DBConnection.Execute SQL
g_Settings.DBConnection.CommitTrans
Exit Sub
FINE:
 MsgBox "Creazione degli assegni non riuscita: " & err.Description, vbOKOnly + vbCritical
 g_Settings.DBConnection.RollbackTrans
End Sub
