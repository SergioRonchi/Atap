VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaSaldi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Saldi"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmProvvisoria 
      Height          =   2115
      Left            =   50
      TabIndex        =   5
      Top             =   0
      Width           =   4920
      Begin VB.CommandButton CmdRicercaAnag 
         Caption         =   "&Ricerca Anagrafica"
         Height          =   525
         Left            =   3555
         TabIndex        =   8
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox TxtCodiceAvvocato 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   810
         Width           =   1350
      End
      Begin VB.CommandButton CmdRicercaA 
         Caption         =   "->"
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   810
         Width           =   330
      End
      Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaSaldi.frx":0000
         Caption         =   "StampaSaldi.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaSaldi.frx":0184
         Keys            =   "StampaSaldi.frx":01A2
         Spin            =   "StampaSaldi.frx":0200
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
         TabIndex        =   10
         Tag             =   "necessario Data Registrazione"
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "StampaSaldi.frx":0228
         Caption         =   "StampaSaldi.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaSaldi.frx":03AC
         Keys            =   "StampaSaldi.frx":03CA
         Spin            =   "StampaSaldi.frx":0428
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
         TabIndex        =   15
         Top             =   1560
         Width           =   4545
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod. Cassetta:"
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label LblDescr 
         Caption         =   "Descrizione:"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   120
         Width           =   825
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   3480
      TabIndex        =   4
      Top             =   2880
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   690
      Left            =   50
      TabIndex        =   0
      Top             =   2115
      Width           =   4920
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Normale"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Saldi Precedenti Negativi"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   225
         Width           =   1815
      End
   End
End
Attribute VB_Name = "StampaSaldi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
    
    If IsPrtTableLocked("PrtSaldi") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If

    LockPrtTable ("PrtSaldi")
    
    createQuery
    If Not GetADORecordset("PrtSaldi", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
            Call Stampa.gestioneReport("PrtSaldi", "", 0, crptToWindow, "Saldi.rpt", 1)
       Else
        MsgBox "Nessun dato in stampa per Saldi!", vbInformation, "Attenzione"
    End If
    DelockPrtTable ("PrtSaldi")
    
End Sub


Private Sub createQuery()

Dim qry1, qry2 As String



Dim qry As String
Dim visitato As Boolean

    qry1 = ""
    qry2 = ""
    qrySQL = " SELECT Saldi.codice, AnagraficaAvvocati.NOME, " & _
             "MID(Chiusura,7,2) & '/' & MID(Chiusura,5,2) & '/' & MID(Chiusura,1,4) , " & _
             "Saldi.Commento,Saldi.numOrdinamento,'" & IIf(OptTipoStampa(1).value, "N", "P") & "'," & _
             "Saldi.SaldoAdempEuro,Saldi.SaldoSfpgEuro,Saldi.SaldoNotifEuro,Saldi.SaldoDecrIngEuro," & _
             "Saldi.SaldoTotaleEuro,'E' "
    qrySQL = qrySQL & " FROM Saldi INNER JOIN AnagraficaAvvocati ON Saldi.Codice = AnagraficaAvvocati.CODAVV "

    visitato = False
    
    If Trim(TxtCodiceAvvocato.Text) <> "" Then
        qry1 = "(Saldi.Codice = '" & TxtCodiceAvvocato.Text & "')"
    End If
    
    If Trim(TxtRicDataIn.Text <> "") And Trim(TxtRicDataFin.Text <> "") Then
        qry2 = "( Saldi.Chiusura >= '" & Format(TxtRicDataIn.Text, "yyyymmdd") & "')"
        qry2 = qry2 & " AND ( Saldi.Chiusura <= '" & Format(TxtRicDataFin, "yyyymmdd") & "')"
    Else
        If Trim(TxtRicDataIn.Text <> "") And Trim(TxtRicDataFin.Text = "") Then
            qry2 = "( Saldi.Chiusura >= '" & Format(TxtRicDataIn.Text, "yyyymmdd") & "')"
        End If
        If Trim(TxtRicDataIn.Text = "") And Trim(TxtRicDataFin.Text <> "") Then
            qry2 = "( Saldi.Chiusura <= '" & Format(TxtRicDataFin.Text, "yyyymmdd") & "')"
        End If
    End If
              
    If qry1 <> "" And qry2 <> "" Then
        qrySQL = qrySQL & " WHERE " & qry1 & " AND " & qry2
        visitato = True
    End If
    If qry1 = "" And qry2 <> "" Then
        qrySQL = qrySQL & " WHERE " & qry2
        visitato = True
    End If
    If qry1 <> "" And qry2 = "" Then
        qrySQL = qrySQL & " WHERE " & qry1
        visitato = True
    End If
    
    If visitato = False Then
        If OptTipoStampa(1).value = True Then
                qrySQL = qrySQL & " WHERE  (((Saldi.SaldoTotaleEuro)<=-" & Str(g_Settings.LimiteSaldo) & "))"
        Else
                qrySQL = qrySQL & " WHERE (((Saldi.SaldoTotaleEuro)>-" & Str(g_Settings.LimiteSaldo) & " And (Saldi.SaldoTotaleEuro)<" & Str(g_Settings.LimiteSaldo) & " And (Saldi.SaldoTotaleEuro)<>0))"
        End If
    Else
        If OptTipoStampa(1).value = True Then
                qrySQL = qrySQL & " AND  (((Saldi.SaldoTotaleEuro)<=-" & Str(g_Settings.LimiteSaldo) & "))"

        Else
                qrySQL = qrySQL & " AND (((Saldi.SaldoTotaleEuro)>-" & Str(g_Settings.LimiteSaldo) & " And (Saldi.SaldoTotaleEuro)<" & Str(g_Settings.LimiteSaldo) & " And (Saldi.SaldoTotaleEuro)<>0))"

        End If
    End If
    
    OpenProgress ("Attendere... Preparazione Stampa!")
    
    qry = "DELETE * FROM PrtSaldi;"
    g_Settings.DBConnection.Execute qry
    qry = "INSERT INTO PRTSALDI (Codice,NOME,CHIUSURA,COMMENTO,NUMORDINAMENTO," & _
          "TIPOSTAMPA,SaldoAdemp,SaldoSfpg,SaldoNotif,SaldoDecrIng,SaldoTotale,Valuta) " & _
          qrySQL
    g_Settings.DBConnection.Execute qry
    UpdateProgress (100)
    CloseProgress
    
End Sub

Private Sub Form_Load()
   Set moFilterManager = New CFilterManager
   moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv
   
  
   Set moFilterManager = New CFilterManager
End Sub
Private Sub moFilterManager_Validate(IsValid As Boolean)
   cmdOk.Enabled = IsValid
End Sub
