VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaGiornalieraAttivita 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Giornaliera Attivita Ricevute"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   2880
      TabIndex        =   10
      Top             =   3360
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   3210
      Left            =   50
      TabIndex        =   5
      Top             =   0
      Width           =   4245
      Begin VB.Frame FrmMetodoStampa 
         Caption         =   "Modalità Stampa"
         Height          =   645
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   4065
         Begin VB.OptionButton OptModSt 
            Caption         =   "Diretta"
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   8
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton OptModSt 
            Caption         =   "Anteprima"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1680
         End
      End
      Begin VB.Frame FrmAttivita 
         Caption         =   "Attività"
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4020
         Begin VB.CheckBox chkSfrattiUNEP 
            Caption         =   "Sfratti UNEP"
            Height          =   330
            Left            =   2160
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox ChkNotificheUENP 
            Caption         =   "Notifiche UNEP"
            Height          =   285
            Left            =   2160
            TabIndex        =   15
            Top             =   240
            Width           =   1620
         End
         Begin VB.CheckBox ChkDecretiIng 
            Caption         =   "Decreti Ingiuntivi"
            Height          =   285
            Left            =   225
            TabIndex        =   3
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox ChkSfrattiPig 
            Caption         =   "Sfratti/Pignoramenti"
            Height          =   330
            Left            =   225
            TabIndex        =   2
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox ChkNotifiche 
            Caption         =   "Notifiche"
            Height          =   285
            Left            =   225
            TabIndex        =   1
            Top             =   1320
            Width           =   1140
         End
         Begin VB.CheckBox ChkAdempimenti 
            Caption         =   "Adempimenti"
            Height          =   195
            Left            =   225
            TabIndex        =   0
            Top             =   315
            Width           =   1455
         End
      End
      Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaGiornalieraAttivita.frx":0000
         Caption         =   "StampaGiornalieraAttivita.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaGiornalieraAttivita.frx":0184
         Keys            =   "StampaGiornalieraAttivita.frx":01A2
         Spin            =   "StampaGiornalieraAttivita.frx":0200
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
         TabIndex        =   12
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaGiornalieraAttivita.frx":0228
         Caption         =   "StampaGiornalieraAttivita.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaGiornalieraAttivita.frx":03AC
         Keys            =   "StampaGiornalieraAttivita.frx":03CA
         Spin            =   "StampaGiornalieraAttivita.frx":0428
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
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   120
         Width           =   870
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "StampaGiornalieraAttivita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qrySQLSfrPig As String
Dim qrySQLNotif As String
Dim qrySQLDec As String
Private WithEvents m_PrintmanagerAdempi As Stampa
Attribute m_PrintmanagerAdempi.VB_VarHelpID = -1
Private WithEvents m_PrintmanagerDecreti As Stampa
Attribute m_PrintmanagerDecreti.VB_VarHelpID = -1
Private WithEvents m_PrintmanagerSfratti As Stampa
Attribute m_PrintmanagerSfratti.VB_VarHelpID = -1
Private WithEvents m_PrintmanagerNotifiche As Stampa
Attribute m_PrintmanagerNotifiche.VB_VarHelpID = -1
Private WithEvents m_prtGiornalieraNotificheUNEP As Stampa
Attribute m_prtGiornalieraNotificheUNEP.VB_VarHelpID = -1
Private WithEvents m_PrtGiornalieraSfrattiPigUNEP As Stampa
Attribute m_PrtGiornalieraSfrattiPigUNEP.VB_VarHelpID = -1




Private Sub CmdAnnulla_Click()
Unload Me
End Sub

Private Sub CmdOK_Click()
Dim frm As Form
   
    

    
     Dim msg As String
    OpenProgress ("Stampa Giornaliera ...")
    
    'ADEMPIMENTI
    If ChkAdempimenti.value = Checked Then
       If Not IsPrtTableLocked("PrtGiornalieraAdempimenti") Then
           LockPrtTable ("PrtGiornalieraAdempimenti")
           Riempi_PRT_GiornalieraAdempimenti
           If Not GetADORecordset("PrtGiornalieraAdempimenti", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
               Set m_PrintmanagerAdempi = New Stampa
              Call m_PrintmanagerAdempi.gestioneReport("PrtGiornalieraAdempimenti", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraAdempimenti.rpt", 1)
            Else
              msg = msg & "Nessun Adempimento di Cancelleria;" & vbCrLf
           End If
        Else
           msg = msg & "Adempimenti bloccati da un altro utente.;" & vbCrLf
       End If
    End If
    
    UpdateProgress (20)
    'DECRETI
    If ChkDecretiIng.value = Checked Then
     If Not IsPrtTableLocked("PrtGiornalieraDecretiIngiuntivi") Then
           LockPrtTable ("PrtGiornalieraDecretiIngiuntivi")
           Riempi_PRT_GiornalieraDecretiIngiuntivi
           If Not GetADORecordset("PrtGiornalieraDecretiIngiuntivi", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
           
              Set m_PrintmanagerDecreti = New Stampa
              Call m_PrintmanagerDecreti.gestioneReport("PrtGiornalieraDecretiIngiuntivi", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraDecreti.rpt", 1)
            Else
              msg = msg & "Nessun Decreto ingiuntivo;" & vbCrLf
           End If
        Else
           msg = msg & "Decreti bloccati da un altro utente.;" & vbCrLf
       End If
    End If
    
    
    UpdateProgress (40)
    'NOTIFICHE
    If ChkNotifiche.value = Checked Then
     If Not IsPrtTableLocked("PrtGiornalieraNotifiche") Then
           LockPrtTable ("PrtGiornalieraNotifiche")
           Riempi_PRT_GiornalieraNotifiche (False)
           If Not GetADORecordset("PrtGiornalieraNotifiche", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
              
              Set m_PrintmanagerNotifiche = New Stampa
              Call m_PrintmanagerNotifiche.gestioneReport("PrtGiornalieraNotifiche", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraNotifiche.rpt", 1)
            Else
              msg = msg & "Nessuna Notifica;" & vbCrLf
           End If
        Else
           msg = msg & "Notifiche bloccate da un altro utente.;" & vbCrLf
       End If
    End If
    UpdateProgress (65)
    'NOTIFICHE UNEP
    If ChkNotificheUENP = Checked Then
     If Not IsPrtTableLocked("PrtGiornalieraNotificheUNEP") Then
           LockPrtTable ("PrtGiornalieraNotificheUNEP")
           Riempi_PRT_GiornalieraNotifiche (True)
           If Not GetADORecordset("PrtGiornalieraNotificheUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
              
              Set m_prtGiornalieraNotificheUNEP = New Stampa
              Call m_prtGiornalieraNotificheUNEP.gestioneReport("PrtGiornalieraNotificheUNEP", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraNotificheUNEP.rpt", 1)
            Else
              msg = msg & "Nessuna Notifica UNEP;" & vbCrLf
           End If
        Else
           msg = msg & "Notifiche UNEP bloccate da un altro utente.;" & vbCrLf
       End If
    End If
    
    UpdateProgress (75)
    'SFRATTI
    If ChkSfrattiPig.value = Checked Then
     If Not IsPrtTableLocked("PrtGiornalieraSfrattiPig") Then
           LockPrtTable ("PrtGiornalieraSfrattiPig")
           Riempi_PRT_GiornalieraSfrattiPignoramenti (False)
           If Not GetADORecordset("PrtGiornalieraSfrattiPig", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
              
              Set m_PrintmanagerSfratti = New Stampa
              Call m_PrintmanagerSfratti.gestioneReport("PrtGiornalieraSfrattiPig", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraSfrattiPig.rpt", 1)
            Else
              msg = msg & "Nessuno Sfratto;" & vbCrLf
           End If
        Else
           msg = msg & "Sfratti bloccati da un altro utente.;" & vbCrLf
       End If
    End If
    
    UpdateProgress (90)
    'SFRATTI UNEP
    If chkSfrattiUNEP.value = Checked Then
     If Not IsPrtTableLocked("PrtGiornalieraSfrattiPigUNEP") Then
           LockPrtTable ("PrtGiornalieraSfrattiPigUNEP")
           Riempi_PRT_GiornalieraSfrattiPignoramenti (True)
           If Not GetADORecordset("PrtGiornalieraSfrattiPigUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
              
              Set m_PrtGiornalieraSfrattiPigUNEP = New Stampa
              Call m_PrtGiornalieraSfrattiPigUNEP.gestioneReport("PrtGiornalieraSfrattiPigUNEP", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "GiornalieraSfrattiPigUNEP.rpt", 1)
            Else
              msg = msg & "Nessuno Sfratto UNEP" & vbCrLf
           End If
        Else
           msg = msg & "Sfratti bloccati da un altro utente.;" & vbCrLf
       End If
    End If
    
    UpdateProgress (100)
    CloseProgress
   
    
    If msg <> "" Then MsgBox msg, vbInformation
    
    DelockPrtTable ("PrtGiornalieraAdempimenti")
    DelockPrtTable ("PrtGiornalieraNotifiche")
    DelockPrtTable ("PrtGiornalieraDecretiIngiuntivi")
    DelockPrtTable ("PrtGiornalieraSfrattiPig")
    
    
End Sub


Private Sub Form_Load()
    TxtRicDataIn = Date
    TxtRicDataFin = Date
    
End Sub

Public Sub Riempi_PRT_GiornalieraAdempimenti()

Dim qrySQL As String
Dim qry As String
Dim qry1, qry2 As String
Dim NumErrori As Integer
    
On Error GoTo Riempi_PRT_GiornalieraAdempimenti

    qry1 = ""
    qry2 = ""
    If TxtRicDataIn.Text <> "" Then
       qry1 = "( DataRegistrazione >= '" & Format(TxtRicDataIn.Text, "YYYYMMDD") & "')"
    End If
    If TxtRicDataFin.Text <> "" Then
        qry2 = "( DataRegistrazione <= '" & Format(TxtRicDataFin.Text, "YYYYMMDD") & "')"
    End If
    
    qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale, ADEMPI.CODAVV, ADEMPI.Progressivo, ADEMPI.NumOrdinamento, "
    qrySQL = qrySQL & "MID (DataRegistrazione,7,2) & '/' & MID (DataRegistrazione,5,2) & '/' & MID (DataRegistrazione,1,4), " & _
                      "ADEMPI.AttivitaRichiesta, ADEMPI.ImpSaldoE,'" & TxtRicDataIn.Text & "','" & TxtRicDataFin.Text & "','E' "
    qrySQL = qrySQL & "FROM ADEMPI RIGHT JOIN TribunaliAppartenenza ON ADEMPI.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale "
    qrySQL = qrySQL & "WHERE ((ADEMPI.Annullo) = 'V'  )"
    
    If qry1 <> "" And qry2 <> "" Then
        qrySQL = qrySQL & " AND " & qry1 & " AND " & qry2
    End If
    If qry1 = "" And qry2 <> "" Then
            qrySQL = qrySQL & " AND " & qry2
    End If
    If qry1 <> "" And qry2 = "" Then
            qrySQL = qrySQL & " AND " & qry1
    End If



qry = "DELETE  * FROM PrtGiornalieraAdempimenti;"
g_Settings.DBConnection.Execute qry

qry = "INSERT INTO PrtGiornalieraAdempimenti (DescrizioneTribunale,codAvv, " & _
      "Progressivo,numOrdinamento,DataRegistrazione,AttivitaRichiesta,ImpSaldo," & _
      "DATA_INIZIO,DATA_FINE,Valuta) " & qrySQL
      
g_Settings.DBConnection.Execute qry

Exit Sub

Riempi_PRT_GiornalieraAdempimenti:
 MsgBox err.Description & vbCrLf & qry
End Sub

Public Sub Riempi_PRT_GiornalieraNotifiche(isUnep As Boolean)

Dim qrySQL As String
Dim qry As String
Dim qry1, qry2 As String
Dim NumErrori As Integer
Dim table As String
On Error GoTo Riempi_PRT_GiornalieraNotifiche
table = "Notifiche"
If isUnep Then table = table & "_UNEP"

    qry1 = ""
    qry2 = ""
    If TxtRicDataIn.Text <> "" Then
       qry1 = "( DataRegistrazione >= '" & ElaboraData(Replace(TxtRicDataIn.Text, "/", "")) & "')"
    End If
    If TxtRicDataFin.Text <> "" Then
        qry2 = "( DataRegistrazione <= '" & ElaboraData(Replace(TxtRicDataFin.Text, "/", "")) & "')"
    End If
    
     If isUnep Then
      qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale, Crono,"
    Else
      qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale,"
    End If
    
    qrySQL = qrySQL & " " & table & ".CODAVV, " & table & ".NumeroAtto," & _
             "Localita1, Note, "
    qrySQL = qrySQL & "MID (DataRegistrazione,7,2) & '/' & MID (DataRegistrazione,5,2) & '/' & MID (DataRegistrazione,1,4),  " & _
             "Parte1, Parte2, TipoAtto.Codice, " & _
             "MID (DataPresentazione,7,2) & '/' & MID (DataPresentazione,5,2) & '/' & MID (DataPresentazione,1,4), "
    qrySQL = qrySQL & "ImpSaldoE, " & table & ".NumOrdinamento, " & _
             "MID (DataRestituzione,7,2) & '/' & MID (DataRestituzione,5,2) & '/' & MID (DataRestituzione,1,4), " & _
             "MID (DataNotifica,7,2) & '/' & MID (DataNotifica,5,2) & '/' & MID (DataNotifica,1,4), " & _
             "'" & TxtRicDataIn.Text & "','" & TxtRicDataFin.Text & "','E' "
             
    qrySQL = qrySQL & "FROM (" & table & " INNER JOIN TribunaliAppartenenza ON " & table & ".CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) "
    qrySQL = qrySQL & "INNER JOIN TipoAtto ON " & table & ".CodAtto = TipoAtto.Codice "
    qrySQL = qrySQL & "WHERE " & table & ".Annullo = 'V'  "
    
  
    
    If qry1 <> "" And qry2 <> "" Then
        qrySQL = qrySQL & " AND " & qry1 & " AND " & qry2
    End If
    If qry1 = "" And qry2 <> "" Then
            qrySQL = qrySQL & " AND " & qry2
    End If
    If qry1 <> "" And qry2 = "" Then
            qrySQL = qrySQL & " AND " & qry1
    End If







 If isUnep Then
      qry = "DELETE * FROM PrtGiornalieraNotificheUNEP"
      g_Settings.DBConnection.Execute qry
      
      qry = "INSERT INTO PrtGiornalieraNotificheUNEP (DescrizioneTribunale,crono, codAvv, "
    Else
      qry = "DELETE * FROM PrtGiornalieraNotifiche;"
      g_Settings.DBConnection.Execute qry
      qry = "INSERT INTO PrtGiornalieraNotifiche (DescrizioneTribunale,codAvv, "
    End If

qry = qry & "NumeroAtto,localita1,[Note],DataRegistrazione,Parte1,Parte2,Descrizione,DataPresentazione, " & _
      "ImpSaldo,numOrdinamento,DataRestituzione,DataNotifica,DATA_INIZIO,DATA_FINE,Valuta) " & qrySQL
g_Settings.DBConnection.Execute qry
Exit Sub

Riempi_PRT_GiornalieraNotifiche:
 MsgBox err.Description & vbCrLf & qry
End Sub

Public Sub Riempi_PRT_GiornalieraSfrattiPignoramenti(isUnep As Boolean)

Dim qrySQL As String
Dim qry As String
Dim qry1, qry2 As String
Dim NumErrori As Integer
Dim table As String
table = "SFRATTI"
If isUnep Then table = table & "_UNEP"

On Error GoTo Riempi_PRT_GiornalieraSfrattiPignoramenti

    qry1 = ""
    qry2 = ""
    If TxtRicDataIn.Text <> "" Then
       qry1 = "( DataRegistrazione >= '" & Format(TxtRicDataIn.Text, "yyyymmdd") & "')"
    End If
    If TxtRicDataFin.Text <> "" Then
        qry2 = "( DataRegistrazione <= '" & Format(TxtRicDataFin.Text, "yyyymmdd") & "')"
    End If
    
    If isUnep Then
      qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale, Crono,"
    Else
      qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale,"
    End If
    
    qrySQL = qrySQL & " " & table & ".CODAVV, " & table & ".NumeroAtto, "
    qrySQL = qrySQL & "MID (DataRegistrazione,7,2) & '/' & MID (DataRegistrazione,5,2) & '/' & MID (DataRegistrazione,1,4), " & _
                      "MID (DataPresentazione,7,2) & '/' & MID (DataPresentazione,5,2) & '/' & MID (DataPresentazione,1,4), " & _
                      "MID (DataRestituzione,7,2) & '/' & MID (DataRestituzione,5,2) & '/' & MID (DataRestituzione,1,4), "
    qrySQL = qrySQL & "Parte1, Parte2, Pignoramenti.Descrizione,  "
    qrySQL = qrySQL & "ImpSaldoE, " & table & ".NumOrdinamento, Localita1, " & _
                      "'" & TxtRicDataIn.Text & "','" & TxtRicDataFin.Text & "','E' "
    qrySQL = qrySQL & "FROM (" & table & " INNER JOIN TribunaliAppartenenza ON "
    qrySQL = qrySQL & "" & table & ".CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) INNER JOIN "
    qrySQL = qrySQL & "Pignoramenti ON " & table & ".CodicePignoramenti = Pignoramenti.Codice WHERE " & table & ".Annullo='V' "
    
    
    If qry1 <> "" And qry2 <> "" Then
        qrySQL = qrySQL & " AND " & qry1 & " AND " & qry2
    End If
    If qry1 = "" And qry2 <> "" Then
            qrySQL = qrySQL & " AND " & qry2
    End If
    If qry1 <> "" And qry2 = "" Then
            qrySQL = qrySQL & " AND " & qry1
    End If



 If isUnep Then
      qry = "DELETE * FROM PrtGiornalieraSfrattiPigUNEP"
      g_Settings.DBConnection.Execute qry
      
      qry = "INSERT INTO PrtGiornalieraSfrattiPigUNEP (DescrizioneTribunale,crono, codAvv, "
    Else
      qry = "DELETE * FROM PrtGiornalieraSfrattiPig;"
      g_Settings.DBConnection.Execute qry
      qry = "INSERT INTO PrtGiornalieraSfrattiPig (DescrizioneTribunale,codAvv, "
    End If



qry = qry & "NumeroAtto,DataRegistrazione,DataPresentazione,DataRestituzione,Parte1,Parte2,Descrizione, " & _
      "ImpSaldo,numOrdinamento, localita1,DATA_INIZIO,DATA_FINE,Valuta) " & qrySQL
g_Settings.DBConnection.Execute qry

Exit Sub

Riempi_PRT_GiornalieraSfrattiPignoramenti:
 MsgBox err.Description & vbCrLf & qry
End Sub

Public Sub Riempi_PRT_GiornalieraDecretiIngiuntivi()

Dim qrySQL As String
Dim qry As String
Dim qry1, qry2 As String
Dim NumErrori As Integer
    
On Error GoTo Riempi_PRT_GiornalieraDecretiIngiuntivi

    qry1 = ""
    qry2 = ""
    If TxtRicDataIn.Text <> "" Then
       qry1 = "( DataRegistrazione >= '" & ElaboraData(Replace(TxtRicDataIn.Text, "/", "")) & "')"
    End If
    If TxtRicDataFin.Text <> "" Then
        qry2 = "( DataRegistrazione <= '" & ElaboraData(Replace(TxtRicDataFin.Text, "/", "")) & "')"
    End If
    
    qrySQL = "SELECT TribunaliAppartenenza.DescrizioneTribunale, DecretiIngiuntivi.CODAVV, "
    qrySQL = qrySQL & "DecretiIngiuntivi.NumeroDecreto, " & _
                      "MID (DataRegistrazione,7,2) & '/' & MID (DataRegistrazione,5,2) & '/' & MID (DataRegistrazione,1,4), " & _
                      "DecretiIngiuntivi.FormulaEsec, "
    qrySQL = qrySQL & "DecretiIngiuntivi.Ricorrente, DecretiIngiuntivi.Debitore, Autorita.Descrizione, "
    qrySQL = qrySQL & "DecretiIngiuntivi.Esenzione, DecretiIngiuntivi.NumeroIngiunzione,  "
    qrySQL = qrySQL & "DecretiIngiuntivi.ImpSaldoE,  DecretiIngiuntivi.NumOrdinamento, " & _
                      "'" & TxtRicDataIn.Text & "','" & TxtRicDataFin.Text & "','E' "
    qrySQL = qrySQL & "FROM (DecretiIngiuntivi INNER JOIN TribunaliAppartenenza ON "
    qrySQL = qrySQL & "DecretiIngiuntivi.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) INNER JOIN Autorita "
    qrySQL = qrySQL & "ON DecretiIngiuntivi.CodAutorita = Autorita.Codice WHERE ((DecretiIngiuntivi.Annullo)='V')"
    
    If qry1 <> "" And qry2 <> "" Then
        qrySQL = qrySQL & " AND " & qry1 & " AND " & qry2
    End If
    If qry1 = "" And qry2 <> "" Then
            qrySQL = qrySQL & " AND " & qry2
    End If
    If qry1 <> "" And qry2 = "" Then
            qrySQL = qrySQL & " AND " & qry1
    End If



qry = "DELETE * FROM PrtGiornalieraDecretiIngiuntivi;"
g_Settings.DBConnection.Execute qry

qry = "INSERT INTO PrtGiornalieraDecretiIngiuntivi (DescrizioneTribunale,codAvv, " & _
      "NumeroDecreto,DataRegistrazione,FormulaEsec,Ricorrente,Debitore,Descrizione," & _
      "Esenzione,NumeroIngiunzione, " & _
      "ImpSaldo,numOrdinamento, DATA_INIZIO,DATA_FINE,Valuta) " & qrySQL

g_Settings.DBConnection.Execute qry
Exit Sub

Riempi_PRT_GiornalieraDecretiIngiuntivi:
     MsgBox err.Description & vbCrLf & qry
End Sub

Private Sub m_Printmanager_StampaEseguita(table As String)
    DelockPrtTable (table)
End Sub
