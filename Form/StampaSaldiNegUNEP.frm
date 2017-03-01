VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaSaldiNegativiUNEP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Saldi Provvisori UNEP"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Periodo"
      Height          =   615
      Left            =   180
      TabIndex        =   22
      Top             =   3360
      Width           =   4920
      Begin VB.OptionButton optMese 
         Caption         =   "Mese"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMese 
         Caption         =   "Bimestre"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.PictureBox PictureUNEP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "StampaSaldiNegUNEP.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   21
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   3720
      TabIndex        =   17
      Top             =   4920
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2280
      TabIndex        =   16
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Frame v 
      Caption         =   "Tipo Stampa"
      Height          =   1050
      Left            =   180
      TabIndex        =   12
      Top             =   2160
      Width           =   4920
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Saldi Compresi  tra  € -5,16 ed  € 5,16"
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Top             =   630
         Width           =   4515
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Saldi Positivi"
         Height          =   375
         Index           =   2
         Left            =   3330
         TabIndex        =   14
         Top             =   225
         Width           =   1410
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Saldi Negativi"
         Height          =   375
         Index           =   1
         Left            =   1575
         TabIndex        =   13
         Top             =   225
         Width           =   1410
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Normale"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame FrmProvvisoria 
      Height          =   2115
      Left            =   180
      TabIndex        =   7
      Top             =   0
      Width           =   4920
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   810
         Width           =   330
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Top             =   810
         Width           =   1350
      End
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   525
         Left            =   3555
         TabIndex        =   2
         Top             =   810
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaSaldiNegUNEP.frx":0442
         Caption         =   "StampaSaldiNegUNEP.frx":055A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaSaldiNegUNEP.frx":05C6
         Keys            =   "StampaSaldiNegUNEP.frx":05E4
         Spin            =   "StampaSaldiNegUNEP.frx":0642
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
         Left            =   2400
         TabIndex        =   19
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaSaldiNegUNEP.frx":066A
         Caption         =   "StampaSaldiNegUNEP.frx":0782
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaSaldiNegUNEP.frx":07EE
         Keys            =   "StampaSaldiNegUNEP.frx":080C
         Spin            =   "StampaSaldiNegUNEP.frx":086A
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
         TabIndex        =   20
         Top             =   1560
         Width           =   4545
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   120
         Width           =   825
      End
      Begin VB.Label LblDescr 
         Caption         =   "Descrizione:"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta:"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   840
         Width           =   1110
      End
   End
   Begin VB.Frame FrmMetodoStampa 
      Caption         =   "Modalità Stampa"
      Height          =   645
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4920
      Begin VB.OptionButton OptModSt 
         Caption         =   "Diretta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   270
         Width           =   1680
      End
      Begin VB.OptionButton OptModSt 
         Caption         =   "Anteprima"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1680
      End
   End
End
Attribute VB_Name = "StampaSaldiNegativiUNEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TrasferimentoOK As Boolean
Dim qrySQL As String
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
  Dim avvocatiEstratti As AvvocatiPerEstratto
  Set avvocatiEstratti = GetAvvocatoSingoloPerEstratto(TxtCodiceAvvocato.Text)
    
    If IsPrtTableLocked("PrtSaldiUNEP") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If

    LockPrtTable ("PrtSaldiUNEP")
    
    Riempi_PRT_EstrattoConto
    
    AggiungiAvvocatiQuota TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, IIf(optMese(1).value, g_Settings.QuotaSoci, g_Settings.QuotaSoci / 2)
     

    createQuery
    If Not GetADORecordset("PrtSaldiUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
        Call Stampa.gestioneReport("PrtSaldiUNEP", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "SaldiProvvisoriUNEP.rpt", 1)
       Else
        MsgBox "Nessun dato evaso! Impossibile creare la stampa Saldi Provvisori.", vbInformation, "Attenzione"
    End If
    DelockPrtTable ("PrtSaldiUNEP")
    
    
    
End Sub

Private Sub Form_Load()
    Set moFilterManager = New CFilterManager
    moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv

  
   TxtRicDataIn = getPrecChiusura()
End Sub
Private Sub moFilterManager_Validate(IsValid As Boolean)
   CmdOK.Enabled = IsValid
End Sub

Public Sub Riempi_PRT_EstrattoConto()

'Dim NumEstConto As Integer
Dim qrySQL As String
Dim qryApp As String
Dim qryDelete As String
Dim qry1, qry2, qry3 As String
Dim NumErrori As Integer
    
' Valuta Corrente

    
On Error GoTo Errore_PRT_EstrattoConto

    qry1 = ""
    qry2 = ""
    qry3 = ""
    qryApp = ""
    
    If TxtRicDataIn.Text <> "" Then
       qry1 = " AND ( DataEvasionePratica >= '" & Format(TxtRicDataIn.Text, "yyyymmdd") & "')"
    End If
    If TxtRicDataFin.Text <> "" Then
        qry2 = " AND ( DataEvasionePratica <= '" & Format(TxtRicDataFin, "yyyymmdd") & "')"
    End If
    
    If TxtCodiceAvvocato.Text <> "" Then
        qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & TxtCodiceAvvocato.Text & "')"
    End If
    
    qryApp = qry1 & qry2 & qry3
    
    OpenProgress ("Attendere... Preparazione Stampa!")
    
    'Reset PrtEstrattoConto
    qryDelete = "DELETE * FROM PrtEstrattoContoUNEP;"
    g_Settings.DBConnection.Execute qryDelete
    UpdateProgress (5)
    
    Dim Tabella As String
    
    'Inizio Sfratti
    
    qrySQL = getQrySfratti(IIf(optMese(0).value, 1, 2), True, "Futuro", TxtRicDataIn.Text, TxtRicDataFin.Text, "S") & qryApp & " ORDER BY SFRATTI_UNEP.NumOrdinamento"
    qrySQL = qrySQL & qryApp
    update_EstConto_Sfratti True, "PrtEstrattoContoUNEP", qrySQL ', NumEstConto
    ' Fine Sfratti
    UpdateProgress (45)
    'Inizio Notifiche
    qrySQL = getQryNotifiche(IIf(optMese(0).value, 1, 2), True, "Futuro", TxtRicDataIn.Text, TxtRicDataFin.Text, "S") & qryApp & " ORDER BY Notifiche_UNEP.NumOrdinamento"
    update_EstConto_Notifiche True, "PrtEstrattoContoUNEP", qrySQL ', NumEstConto
    'Fine Notifiche
    UpdateProgress (70)
    UpdateProgress (95)
    
    CloseProgress

Exit Sub

Errore_PRT_EstrattoConto:

   MsgBox "Attenzione errore in stampa Estratto Conto!" & Chr(10) & err & " - " & Error(err), vbCritical, "Attenzione"

End Sub

Private Sub createQuery()
On Error GoTo fine

Dim qry As String
    
    qrySQL = "SELECT PrtEstrattoContoUNEP.CODAVV, PrtEstrattoContoUNEP.Saldo_Precedente, PrtEstrattoContoUNEP.NOME, AnagraficaAvvocati.NumOrdinamento,( Sum(PrtEstrattoContoUNEP.SALDO - PrtEstrattoContoUNEP.QUOTA+ PrtEstrattoContoUNEP.Deduzione)  + PrtEstrattoContoUNEP.Saldo_Precedente ) AS totaleSaldo, " & _
              "'" & TxtRicDataIn.Text & "','" & TxtRicDataFin.Text & "','E' "
    qrySQL = qrySQL & " FROM PrtEstrattoContoUNEP INNER JOIN AnagraficaAvvocati ON PrtEstrattoContoUNEP.CODAVV = AnagraficaAvvocati.CODAVV "
    qrySQL = qrySQL & " GROUP BY PrtEstrattoContoUNEP.CODAVV, PrtEstrattoContoUNEP.Saldo_Precedente, PrtEstrattoContoUNEP.NOME, AnagraficaAvvocati.NumOrdinamento"
    
    If OptTipoStampa(1).value = True Then
       
            qrySQL = qrySQL & " HAVING   (Sum(PrtEstrattoContoUNEP.SALDO - PrtEstrattoContoUNEP.QUOTA + PrtEstrattoContoUNEP.Deduzione)+ PrtEstrattoContoUNEP.Saldo_Precedente)<=-" & Str(g_Settings.LimiteSaldo) & " "
       
    End If
    
    If OptTipoStampa(2).value = True Then
            qrySQL = qrySQL & " HAVING   (Sum(PrtEstrattoContoUNEP.SALDO - PrtEstrattoContoUNEP.QUOTA+ PrtEstrattoContoUNEP.Deduzione)+ PrtEstrattoContoUNEP.Saldo_Precedente)>=" & Str(g_Settings.LimiteSaldo) & ""

    End If
    
    If OptTipoStampa(3).value = True Then
        
            qrySQL = qrySQL & " HAVING  (Sum(PrtEstrattoContoUNEP.SALDO - PrtEstrattoContoUNEP.QUOTA+ PrtEstrattoContoUNEP.Deduzione)+ PrtEstrattoContoUNEP.Saldo_Precedente)>=-" & Str(g_Settings.LimiteSaldo) & " AND (Sum(PrtEstrattoContoUNEP.SALDO - PrtEstrattoContoUNEP.QUOTA+ PrtEstrattoContoUNEP.Deduzione)+ PrtEstrattoContoUNEP.Saldo_Precedente) <=" & Str(g_Settings.LimiteSaldo) & " "
        
    End If

    
    
    
    qry = "DELETE * FROM PrtSaldiUNEP;"
    g_Settings.DBConnection.Execute qry
    
    qry = "INSERT INTO PrtSaldiUNEP (codice,SALDO_PRECEDENTE,NOME," & _
          "numOrdinamento,SaldoTotale,DATA_INIZIO,DATA_FINE,Valuta) " & _
          qrySQL
          
    g_Settings.DBConnection.Execute qry
    Exit Sub
fine:
 MsgBox err.Description & vbCrLf & qry
End Sub

