VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaEstrattoContoUNEP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Estratto Conto UNEP"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictureUNEP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "StampaEstrattoContoUNEP.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   23
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4440
      TabIndex        =   10
      Top             =   5520
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   3000
      TabIndex        =   9
      Top             =   5520
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   5220
      Left            =   0
      TabIndex        =   2
      Top             =   135
      Width           =   5925
      Begin VB.Frame FrmProvvisoria 
         Height          =   2115
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   5640
         Begin VB.Frame fraScelta 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   4695
            Begin VB.CommandButton CmdRicercaA 
               Caption         =   "->"
               Height          =   285
               Left            =   2625
               TabIndex        =   21
               Top             =   0
               Width           =   330
            End
            Begin VB.TextBox TxtCodiceAvvocato 
               Height          =   285
               Left            =   1185
               MaxLength       =   10
               TabIndex        =   20
               Top             =   0
               Width           =   1350
            End
            Begin VB.CommandButton CmdRicercaAnag 
               Caption         =   "&Ricerca Anagrafica"
               Height          =   525
               Left            =   3420
               TabIndex        =   19
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label LblCodAvvocato 
               Caption         =   "Cod. Cassetta:"
               Height          =   255
               Left            =   0
               TabIndex        =   22
               Top             =   30
               Width           =   1110
            End
         End
         Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
            DataField       =   "DataRegistrazione"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoContoUNEP.frx":0442
            Caption         =   "StampaEstrattoContoUNEP.frx":055A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoContoUNEP.frx":05C6
            Keys            =   "StampaEstrattoContoUNEP.frx":05E4
            Spin            =   "StampaEstrattoContoUNEP.frx":0642
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
            TabIndex        =   13
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoContoUNEP.frx":066A
            Caption         =   "StampaEstrattoContoUNEP.frx":0782
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoContoUNEP.frx":07EE
            Keys            =   "StampaEstrattoContoUNEP.frx":080C
            Spin            =   "StampaEstrattoContoUNEP.frx":086A
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
            TabIndex        =   17
            Top             =   1560
            Width           =   4545
         End
         Begin VB.Label LblDescr 
            Caption         =   "Descrizione:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label LblRicDataFin 
            Caption         =   "Data Fine :"
            Height          =   285
            Left            =   2520
            TabIndex        =   15
            Top             =   120
            Width           =   825
         End
         Begin VB.Label LblRicDataIn 
            Caption         =   "Data Inizio :"
            Height          =   285
            Left            =   135
            TabIndex        =   14
            Top             =   120
            Width           =   870
         End
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Sfratti/Pignor."
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   8
         Tag             =   "Pignoramenti"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Notifiche"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Tag             =   "Notifiche"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ChkAbilitaAnteDef 
         Caption         =   "Abilita anteprima in stampa definitiva"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   4800
         Width           =   3015
      End
      Begin VB.Frame FrmMetodoStampa 
         Caption         =   "Modalit� Stampa"
         Height          =   645
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   5640
         Begin VB.OptionButton OptModSt 
            Caption         =   "Anteprima"
            Height          =   195
            Index           =   0
            Left            =   855
            TabIndex        =   1
            Top             =   270
            Value           =   -1  'True
            Width           =   1680
         End
         Begin VB.OptionButton OptModSt 
            Caption         =   "Diretta"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   5
            Top             =   270
            Width           =   1680
         End
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Definitiva"
         Height          =   270
         Index           =   0
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1410
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Provvisoria"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Caption         =   "Periodo"
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   5655
         Begin VB.ComboBox cmbBinestreAnno 
            Height          =   315
            Left            =   4320
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbBimestre 
            Height          =   315
            ItemData        =   "StampaEstrattoContoUNEP.frx":0892
            Left            =   3120
            List            =   "StampaEstrattoContoUNEP.frx":0894
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optMese 
            Caption         =   "Bimestre"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optMese 
            Caption         =   "Mese"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Bimestre:"
            Height          =   255
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "StampaEstrattoContoUNEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Avvocato As String
Private WithEvents moFilterManager As CFilterManager
Attribute moFilterManager.VB_VarHelpID = -1
Public TrasferimentoOK As Boolean

Private Sub ChkAbilitaAnteDef_Click()
    If ChkAbilitaAnteDef.value = 1 Then
        FrmMetodoStampa.Enabled = True
    Else
        FrmMetodoStampa.Enabled = False
        OptModSt(1).value = True
    End If
End Sub

Private Sub CmdAnnulla_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub
Private Sub RisolviOrdinamentoErrato()
 Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
 Dim numOrd As Long
 Dim n As Long
 Dim i As Long
 Set rs = newAdoRs
 
 rs.Open "SELECT NumOrdinamento,Count(*) AS N FROM AnagraficaAvvocati  Group By NumOrdinamento HAVING Count(*)>1", g_Settings.DBConnection
 
 While Not rs.EOF
   numOrd = rs(0)
   n = rs(1)
   
   Set rs1 = newAdoRs
   rs1.Open "SELECT Min(NumOrdinamento) FROM AnagraficaAvvocati  Where NumOrdinamento>" & numOrd, g_Settings.DBConnection
    If Not rs1.EOF Then
      If rs1(0) < numOrd + n Then
        g_Settings.DBConnection.Execute "UPDATE AnagraficaAvvocati SET NumOrdinamento=NumOrdinamento + " & (n - 1) & " WHERE NumOrdinamento>=" & numOrd
      End If
    End If
    rs1.Close
    Set rs2 = newAdoRs
    rs2.Open "SELECT CodAvv FROM AnagraficaAvvocati  Where NumOrdinamento=" & numOrd, g_Settings.DBConnection
    i = 0
    While Not rs2.EOF
      g_Settings.DBConnection.Execute "UPDATE AnagraficaAvvocati SET NumOrdinamento=NumOrdinamento + " & i & " WHERE CodAvv='" & rs2(0) & "'"
      i = i + 1
      rs2.MoveNext
    Wend
    rs2.Close
   rs.MoveNext
 Wend


rs.Close
End Sub
Private Sub CmdOK_Click()
  Dim prov As String
  
  If Not IsDate(TxtRicDataIn.Text) Or Not IsDate(TxtRicDataFin.Text) Then
    MsgBox "Inserire l'intervallo di date", vbOKOnly + vbCritical
    Exit Sub
  End If
  RisolviOrdinamentoErrato
  
  prov = "N"
  If OptTipoStampa(1).value Then prov = "S"
     
    If IsPrtTableLocked("PrtAssegniCircolariUNEP") Or IsPrtTableLocked("PrtEstrattoContoUNEP") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' gi� in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilit�'", vbInformation + vbOKOnly
      Exit Sub
    End If
    LockPrtTable ("PrtAssegniCircolariUNEP")
    LockPrtTable ("PrtEstrattoContoUNEP")

    g_Settings.DBConnection.Execute "DELETE * FROM PrtData"
    
    g_Settings.DBConnection.Execute "INSERT INTO PrtData(Tipo, Bimestre, BimestreAnno) VALUES(" & IIf(optMese(0).value, 1, 2) & "," & cmbBimestre.ListIndex + 1 & "," & cmbBinestreAnno.List(cmbBinestreAnno.ListIndex) & ")"
    
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, TxtCodiceAvvocato.Text, 0, Chk(2), 0, Chk(3), prov, True
           
    If TxtCodiceAvvocato.Text = "" Then AggiungiAvvocatiSenzaOperazioni TxtRicDataIn.Text, TxtRicDataFin.Text, TxtCodiceAvvocato.Text
    
    If Not GetADORecordset("PrtEstrattoContoUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
        If OptTipoStampa(0).value = True Or ChkAbilitaAnteDef.value = True Then
            GestStampaDefinitiva
        Else
      
          Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "EstrattoContoUNEP.rpt", 1, "Tipo='ESTRATTO'")
          If Stampa.Destination = crptToPrinter Then
                    Unload Stampa
          End If
        End If
    Else
        MsgBox "Nessun dato evaso! Impossibile creare l'Estratto Conto.", vbInformation, "Attenzione"
      
    End If
    
    
    DelockPrtTable ("PrtAssegniCircolariUNEP")
    DelockPrtTable ("PrtEstrattoContoUNEP")
  
End Sub



Private Sub Form_Load()
Dim c As Control
    Set moFilterManager = New CFilterManager
    moFilterManager.Initialize TxtRicDataIn, TxtRicDataFin, TxtCodiceAvvocato, CmdRicercaA, CmdRicercaAnag, LblDescrCodAvv

    ChkAbilitaAnteDef.Enabled = False
    Me.Move 400, 400
    For Each c In Chk
      c.value = GetSetting("ATAP", "Config", c.Tag, 1)
    Next
    m_Avvocato = "ALL"
    
     optMese(1).value = True
     optMese_Click (1)
    
End Sub
Private Sub moFilterManager_Validate(IsValid As Boolean)
   CmdOK.Enabled = IsValid
End Sub

Private Sub optMese_Click(Index As Integer)
 Dim Y As Integer, currentYear As Integer, currentMonth As Integer
    Dim currentBimestre As Integer
    Dim precBimestre As Integer
    Dim precMese As Integer
    
    currentMonth = month(Now)
    currentYear = year(Now)
    currentBimestre = currentMonth \ 2
    precBimestre = currentBimestre - 1
    precMese = currentMonth - 1
    
    
    
 cmbBimestre.Clear
 cmbBinestreAnno.Clear
 
     For Y = currentYear - 1 To currentYear + 10
      cmbBinestreAnno.AddItem Y
    Next
 
 If Index = 0 Then
   Label1.Caption = "Mese"
   cmbBimestre.AddItem ("Gennaio")
   cmbBimestre.AddItem ("Febbraio")
   cmbBimestre.AddItem ("Marzo")
   cmbBimestre.AddItem ("Aprile")
   cmbBimestre.AddItem ("Maggio")
   cmbBimestre.AddItem ("Giugno")
   cmbBimestre.AddItem ("Luglio")
   cmbBimestre.AddItem ("Agosto")
   cmbBimestre.AddItem ("Settembre")
   cmbBimestre.AddItem ("Ottobre")
   cmbBimestre.AddItem ("Novembre")
   cmbBimestre.AddItem ("Dicembre")

    
    If precMese < 1 Then
       cmbBinestreAnno.ListIndex = 0
       cmbBimestre.ListIndex = 11
     Else
        cmbBinestreAnno.ListIndex = 1
        cmbBimestre.ListIndex = precMese
    End If
 Else
   Label1.Caption = "Bimestre"
   cmbBimestre.AddItem ("Gen-Feb")
   cmbBimestre.AddItem ("Mar-Apr")
   cmbBimestre.AddItem ("Mag-Giu")
   cmbBimestre.AddItem ("Lug-Ago")
   cmbBimestre.AddItem ("Set-Ott")
   cmbBimestre.AddItem ("Nov-Dic")
    If precBimestre < 0 Then
       cmbBinestreAnno.ListIndex = 0
       cmbBimestre.ListIndex = 5
     Else
        cmbBinestreAnno.ListIndex = 1
        cmbBimestre.ListIndex = precBimestre
    End If
 End If
 

    

End Sub

Private Sub OptTipoStampa_Click(Index As Integer)
    fraScelta.Visible = (OptTipoStampa(1).value = True)
    FrmProvvisoria.Enabled = (OptTipoStampa(1).value = True)
    ChkAbilitaAnteDef.Enabled = Not (OptTipoStampa(1).value = True)
    FrmMetodoStampa.Enabled = (OptTipoStampa(1).value = True)
    OptModSt(1).value = Not (OptTipoStampa(1).value = True)
    OptModSt(0).value = (OptTipoStampa(1).value = True)
    TxtRicDataIn.Enabled = (OptTipoStampa(1).value = True)
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim c As Control


    For Each c In Chk
       SaveSetting "ATAP", "Config", c.Tag, c.value
    Next

End Sub


Public Sub TrasferimentoDatiAlDbStorico()
On Error GoTo ErroreTrasferimento
Dim i As Integer
Dim schema As String
Dim nome As String


  If Chk(2).value = 1 Then schema = schema + Left(Chk(2).Caption, 1)
  If Chk(3).value = 1 Then schema = schema + Left(Chk(3).Caption, 1)


nome = TxtCodiceAvvocato.Text
If nome = "" Then nome = "COMPLETO"
Dim d1 As String
Dim d2 As String



If IsDate(TxtRicDataIn.Text) Then
  d1 = Format(TxtRicDataIn.Text, "yyyymmdd")
Else
  d1 = "19990101"
End If

If IsDate(TxtRicDataFin.Text) Then
  d2 = Format(TxtRicDataFin.Text, "yyyymmdd")
Else
  d2 = Format(Now, "yyyymmdd")
End If


TrasferimentoOK = Trasferisci(g_Settings.StoricoEC_UNEP & "\ECUNEP_" & Format(Date, "yyyymmdd") & "_" & nome & ".mdb", d1, d2, True, Trim(TxtCodiceAvvocato.Text), schema)
 
Exit Sub

ErroreTrasferimento:
        
    MsgBox "Errore durante trasferimento dati nel db storico!", vbInformation, "Attenzione"
    TrasferimentoOK = False
    Exit Sub
    
End Sub




Public Sub UpdateDataUltimoEstConto()

g_Settings.DBConnection.Execute ("UPDATE Date_EstrattiConto SET DATA_ULTIMO_ESTCONTO_UNEP='" & Format(TxtRicDataFin.Text, "dd/mm/yyyy") & "'")
    
End Sub
Public Sub AggiornaSaldo(rsAssegni As ADODB.Recordset)
On Error GoTo FINE
Dim saldo As Double
Dim saldoPrec As Double
Dim dataEC As String
Dim codice As String
Dim SQL As String
Dim Commento As String
Dim prog As String
Dim rs As ADODB.Recordset

 dataEC = TxtRicDataFin.Text
 saldo = rsAssegni!saldo + rsAssegni!SALDO_PRECEDENTE
 codice = rsAssegni!codAvv
 If saldo >= g_Settings.LimiteSaldo Then
   saldo = 0
   
   Commento = ""
   Else
   If saldo <= -g_Settings.LimiteSaldo Then
            Commento = "Saldo Negativo"
   End If
 End If
 Set rs = GetADORecordset("SaldiUNEP", "chiusura", "codice='" & codice & "'", g_Settings.DBConnection)
 If rs Is Nothing Then
   'Record inesistente
   SQL = "INSERT INTO SALDIUNEP(codice,Stato,PROG_Saldi,Commento,SaldoAdemp,SaldoSfpg, " & _
         "SaldoNotif,SaldoDecrIng,SaldoAdempEuro,SaldoSfpgEuro,SaldoNotifEuro,SaldoDecrIngEuro," & _
         "SaldoTotale,SaldoTotaleEuro) " & _
         "VALUES ('" & codice & "','N'," & 1 & ",'" & Commento & "'," & _
         "0,0,0,0,0,0,0,0," & Str(saldo * 1936.27) & "," & Str(saldo) & ");"
 Else
   'record gi� prersente
   If Format(RitornaData(rs!Chiusura), "yyyy") = Format(dataEC, "yyyy") Then
            
            prog = "PROG_Saldi + 1"
          Else
            prog = 1
            
   End If
   SQL = "UPDATE SALDIUNEP SET " & _
         "Stato='N',PROG_Saldi=" & prog & ",Commento='" & Commento & "',SaldoAdemp=0,SaldoSfpg=0, " & _
         "SaldoNotif=0,SaldoDecrIng=0,SaldoAdempEuro=0,SaldoSfpgEuro=0,SaldoNotifEuro=0,SaldoDecrIngEuro=0," & _
         "SaldoTotale=" & Str(saldo * 1936.27) & ",SaldoTotaleEuro=" & Str(saldo) & _
         " WHERE codice='" & codice & "';"
 End If
 g_Settings.DBConnection.Execute SQL
 Exit Sub
FINE:
 MsgBox err.Description & vbCrLf & SQL
 
End Sub

Public Sub AggiornaTabellaSaldi()


Dim rsAssegni As ADODB.Recordset


'CreaTabAppAC_Sal  'Crea AssegniCircolari gi� chiamata prima


Set rsAssegni = GetADORecordset("TempSaldiUNEP", "*", "1=1", g_Settings.DBConnection)
If (Not rsAssegni Is Nothing) Then
  While (Not rsAssegni.EOF)
    AggiornaSaldo rsAssegni
 
    rsAssegni.MoveNext
 
  Wend
End If


'Set GestioneSaldiEstrattoConto = gDB.OpenRecordset("Saldi", dbOpenTable)
'Set TabellaRPT = gDB.OpenRecordset("PrtAssegniCircolari", dbOpenTable)
'GestioneSaldiEstrattoConto.Index = "Codice"
'
'If TabellaRPT.RecordCount > 0 Then
'    TabellaRPT.MoveFirst
'    Do Until TabellaRPT.EOF
'        codice = TabellaRPT!codAvv
'        GestioneSaldiEstrattoConto.Seek "=", codice
'
'        saldo = TabellaRPT!saldo
'        saldoPrec = TabellaRPT!SALDO_PRECEDENTE
'        If Not GestioneSaldiEstrattoConto.NoMatch Then
'            GestioneSaldiEstrattoConto.Edit
'            RiempiTabellaSaldi GestioneSaldiEstrattoConto, dataEC, codice, saldo, saldoPrec
'            GestioneSaldiEstrattoConto.Update
'        Else
'            GestioneSaldiEstrattoConto.AddNew
'            RiempiTabellaSaldi GestioneSaldiEstrattoConto, dataEC, codice, saldo, saldoPrec
'            GestioneSaldiEstrattoConto.Update
'        End If
'        TabellaRPT.MoveNext
'    Loop
'End If
'
'GestioneSaldiEstrattoConto.Close
'TabellaRPT.Close

End Sub
Public Sub aggiornaFattura(ByRef nFat As Long, codice As String, data As String, adempi As Double, _
                            decreti As Double, Notifiche As Double, stratti As Double)
Dim SQL As String
Dim rs As ADODB.Recordset
If codice = "525/158" Then
 Debug.Print "Errore"
End If
Dim quotaBimestrale As Double
quotaBimestrale = GetADOValue("Parametri", "QuotaSoci", "1=1", g_Settings.DBConnection, True)

Dim strBimestre As String
Dim bimestre As Integer
Dim anno As Integer

Dim T As Integer

T = GetADOValue("PrtData", "Tipo", "1=1", g_Settings.DBConnection, True)
bimestre = GetADOValue("PrtData", "Bimestre", "1=1", g_Settings.DBConnection, True)
anno = GetADOValue("PrtData", "BimestreAnno", "1=1", g_Settings.DBConnection, True)

If T = 1 Then quotaBimestrale = quotaBimestrale / 2


Select Case bimestre
  Case 1
    strBimestre = "GENNAIO-FEBBRAIO " & anno
  Case 2
    strBimestre = "MARZO-APRILE " & anno
  Case 3
    strBimestre = "MAGGIO-GIUGNO " & anno
  Case 3
    strBimestre = "LUGLIO-AGOSTO " & anno
  Case 5
    strBimestre = "SETTEMBRE-OTTOBRE " & anno
  Case 6
    strBimestre = "NOVEMBRE-DICEMBRE " & anno
End Select

If GetADORecordset("StoricoFattureUNEP", "*", "codAVV='" & codice & "' and DATAFATTURA='" & Format(data, "yyyymmdd") & "'", g_Settings.DBConnection) Is Nothing Then
     Set rs = GetADORecordset("AnagraficaAvvocati", "*", "codAVV='" & codice & "'", g_Settings.DBConnection)
     
     If rs!AFAT <> "S" Then Exit Sub
     
     SQL = "INSERT INTO StoricoFattureUNEP (numOrdinamento,NOME,INDIRI,LOCALI,PROV,CAP,PIVA,codAvv," & _
           "NumeroFattura,DataFattura,DataFatturaNormale,Valuta,ImportoIva,CompAdempEuro,CompDecrIngEuro,CompNotifEuro,CompSfpgEuro, Bimestre, Quota) " & _
           "VALUES (" & rs!numOrdinamento & ",'" & Replace(Left(rs!nome, 40), "'", "''") & "','" & Replace(Left(rs!INDIRI, 40), "'", "''") & "','" & Replace(Left(rs!LOCALI, 35), "'", "''") & _
           "','" & rs!prov & "','" & rs!CAP & "','" & rs!PIVA & "','" & codice & "'," & nFat & _
           ",'" & Format(data, "yyyymmdd") & "','" & data & "','E',0," & Str(adempi) & "," & Str(decreti) & "," & Str(Notifiche) & "," & Str(stratti) & _
           ",'" & strBimestre & "'," & Str(quotaBimestrale) & ");"
           nFat = nFat + 1
   Else
     SQL = "UPDATE StoricoFattureUNEP SET " & _
           "CompAdempEuro=CompAdempEuro+" & Str(adempi) & _
           ",CompDecrIngEuro=CompDecrIngEuro+" & Str(decreti) & _
           ",CompNotifEuro=CompNotifEuro+" & Str(Notifiche) & _
           ",CompSfpgEuro=CompSfpgEuro+" & Str(stratti) & _
           " WHERE codAVV='" & codice & "' and DATAFATTURA='" & data & "';"
   
End If
g_Settings.DBConnection.Execute SQL

End Sub
Public Sub GeneraFattura(Numero As Integer, data As Date)
Dim nFat As Long
Dim ValEuro As Variant
Dim Query As String
Dim SQL As String
Dim rsEstratto As ADODB.Recordset
Dim codice As String
Dim adempi As Double
Dim decreti As Double
Dim Notifiche As Double
Dim sfratti As Double

ValEuro = 1936.27
nFat = Numero


SQL = "SELECT codAvv,DESCR_ATTIVITA,Sum(Competenze) FROM PrtEstrattoContoUNEP " & _
      "GROUP BY NumOrdinamento,codAvv,DESCR_ATTIVITA " & _
      "ORDER BY NumOrdinamento;"

Set rsEstratto = newAdoRs()
rsEstratto.Open SQL, g_Settings.DBConnection
If rsEstratto.EOF Then Exit Sub

codice = rsEstratto(0)
While Not rsEstratto.EOF
 If rsEstratto(0) = "525/158" Then
   Debug.Print "Strano"
 End If
 If rsEstratto(0) <> codice Then
   
   If adempi + decreti + Notifiche + sfratti > 0 Then
     aggiornaFattura nFat, codice, "" & data, adempi, decreti, Notifiche, sfratti
        codice = rsEstratto(0)
        adempi = 0
        decreti = 0
        Notifiche = 0
        sfratti = 0
      Else
            codice = rsEstratto(0)
            Debug.Print "Importo nullo: " & codice
   End If
   
 End If
  If rsEstratto(1) = "Adempimenti Cancelleria" Then adempi = rsEstratto(2).value
  If rsEstratto(1) = "Decreti Ingiuntivi" Then decreti = rsEstratto(2).value
  If rsEstratto(1) = "Notifiche" Then Notifiche = rsEstratto(2).value
  If rsEstratto(1) = "Sfratti/Pignoramenti" Then sfratti = rsEstratto(2).value
 rsEstratto.MoveNext
Wend
If adempi + decreti + Notifiche + sfratti > 0 Then
     aggiornaFattura nFat, codice, "" & data, adempi, decreti, Notifiche, sfratti
End If
End Sub
Private Sub SalvaSaldiTemporanei()

Dim qry As String

Dim sp1, sp3, sp5 As Double
On Error GoTo FINE


'Reset PrtEstrattoContoUNEP
qry = "DELETE * FROM TempSaldiUNEP;"
g_Settings.DBConnection.Execute qry

qry = GetQuerySaldi("TempSaldiUNEP", " < ")
g_Settings.DBConnection.Execute qry

qry = GetQuerySaldi("TempSaldiUNEP", " >= ")
g_Settings.DBConnection.Execute qry

Exit Sub
FINE:
 MsgBox err.Description

End Sub
Private Sub CreaTabAssegni()

Dim qry As String

Dim sp1, sp3, sp5 As Double
On Error GoTo FINE


'Reset PrtEstrattoContoUNEP
qry = "DELETE * FROM PrtAssegniCircolariUNEP;"
g_Settings.DBConnection.Execute qry
qry = GetQuerySaldi("PrtAssegniCircolariUNEP", " >= ")
g_Settings.DBConnection.Execute qry
Exit Sub
FINE:
 MsgBox err.Description

End Sub
Private Function GetQuerySaldi(destinationTable As String, condition As String) As String
Dim qry As String

qry = "INSERT INTO " & destinationTable & " ( CODAVV, NOME, DESCR_ATTIVITA, DEPOSITO, COMPETENZE, SALDO, " & _
     "SPESE1, SPESE2, SPESE3, SPESE4, SPESE5, SPESE6, SALDO_PRECEDENTE, VALUTA,NumOrdinamento,DATA_INIZIO,DATA_FINE ) " & _
     "SELECT CODAVV, NOME, 'XXX', Sum(PrtEstrattoContoUNEP.DEPOSITO) AS DEP," & _
     "Sum(PrtEstrattoContoUNEP.COMPETENZE)*" & Str(1 + g_Settings.IVA / 100) & " AS [COMP], [DEP]-[COMP]-[S1]-[S2]-[S3]-[S4]-[S5]-[S6] AS Ass," & _
     "Sum(IIF(DESCR_SPESE1='Fotocopie',[SPESE1]*[PrtEstrattoContoUNEP]![QtaFotocopie],[SPESE1])) AS S1, Sum(PrtEstrattoContoUNEP.SPESE2) AS S2," & _
     "Sum(IIF(DESCR_SPESE3='Marche',[SPESE3]*[QtaMarche],[SPESE3])) AS S3, Sum(PrtEstrattoContoUNEP.SPESE4) AS S4, " & _
     "Sum(IIF(DESCR_SPESE5='Diritti Cancelleria',[SPESE5]*[qtaDirittiCancelleria],[SPESE5])) AS S5, Sum(PrtEstrattoContoUNEP.SPESE6) AS S6," & _
     "fIRST(PrtEstrattoContoUNEP.SALDO_PRECEDENTE) AS S_PRECEDENTE, 'E' AS Valuta,NumOrdinamento,DATA_INIZIO,DATA_FINE " & _
     "From PrtEstrattoContoUNEP " & _
     "GROUP BY PrtEstrattoContoUNEP.CODAVV,PrtEstrattoContoUNEP.Saldo_Precedente, PrtEstrattoContoUNEP.NOME, NumOrdinamento,DATA_INIZIO,DATA_FINE " & _
     " HAVING   Sum(PrtEstrattoContoUNEP.DEPOSITO) + fIRST(PrtEstrattoContoUNEP.SALDO_PRECEDENTE) -Sum(PrtEstrattoContoUNEP.COMPETENZE)*" & Str(1 + g_Settings.IVA / 100) & "-" & _
     "Sum(IIF(DESCR_SPESE1='Fotocopie',[SPESE1]*[PrtEstrattoContoUNEP]![QtaFotocopie],[SPESE1]))-" & _
     "Sum(PrtEstrattoContoUNEP.SPESE2) -Sum(IIF(DESCR_SPESE3='Marche',[SPESE3]*[QtaMarche],[SPESE3])) - " & _
     "Sum(PrtEstrattoContoUNEP.SPESE4) - Sum(IIF(DESCR_SPESE5='Diritti Cancelleria',[SPESE5]*[qtaDirittiCancelleria],[SPESE5])) - " & _
     "Sum(PrtEstrattoContoUNEP.SPESE6) " & condition & Str(g_Settings.LimiteSaldo)
  GetQuerySaldi = qry
End Function

Public Sub CreazioneStampaAssegniCircolari()


SalvaSaldiTemporanei

CreaTabAssegni


g_Settings.DBConnection.Execute "DELETE * FROM PrtAssegniCircolariUNEP where (saldo + SALDO_PRECEDENTE)<" & Str(g_Settings.LimiteSaldo)

End Sub

Public Sub GestStampaDefinitiva()

Dim MSG_Avviso, Response As Variant
    MSG_Avviso = "Durante questa operazione � necessario non modificare alcun dato." & Chr(10)
    MSG_Avviso = MSG_Avviso & "Chiudere tutte le finestre aperte e verificare che nessun"
    MSG_Avviso = MSG_Avviso & " altro client abbia l'applicazione attiva!" & Chr(10) & "Proseguire?"
    Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
    If Response = vbYes Then    ' User chose Yes.
            Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "EstrattoContoUNEP.rpt", 2, "Tipo='ESTRATTO'")
            If Stampa.Destination = crptToPrinter Then
                    Unload Stampa
            End If
            While Stampa.IsClosed = False
              DoEvents
              'WAIT
            Wend
            
            MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
            MSG_Avviso = MSG_Avviso & "Si vuole Procedere con la creazione della stampa Richiesta assegni circolari?" & Chr(10)
            MSG_Avviso = MSG_Avviso & "(Obbligatorio per rendere definitivo l'estratto conto)"
            Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
            If Response = vbYes Then    ' User chose Yes.
                CreazioneStampaAssegniCircolari
                Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "AssegniCircolariUNEP.rpt", 3)
                If Stampa.Destination = crptToPrinter Then
                    Unload Stampa
                End If
                While Stampa.IsClosed = False
                    DoEvents
                    'WAIT
                Wend
                
                MSG_Avviso = "Verificare il buon esito della stampa!" & Chr(10)
                MSG_Avviso = MSG_Avviso & "Si vuole Procedere col trasferimento dei dati nel database storico?" & Chr(10)
                MSG_Avviso = MSG_Avviso & "(Obbligatorio per rendere definitivo l'estratto conto)"
                Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
                If Response = vbYes Then    ' User chose Yes.
                    TrasferimentoDatiAlDbStorico
                    
                    If TrasferimentoOK = True Then
                        UpdateDataUltimoEstConto
                        
                        AggiornaTabellaSaldi
                        ImpostazioniFatturazione.isUnep = True
                        ImpostazioniFatturazione.Show
                    End If
               End If
           End If
   End If
End Sub

