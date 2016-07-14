VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaEstrattoConto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Estratto Conto"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4080
      TabIndex        =   12
      Top             =   4920
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   2640
      TabIndex        =   11
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Frame FrmTipoStampa 
      Caption         =   "Tipo Stampa"
      Height          =   4740
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   5325
      Begin VB.Frame FrmProvvisoria 
         Height          =   2115
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   4920
         Begin VB.Frame fraScelta 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   615
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   4695
            Begin VB.CommandButton CmdRicercaA 
               Caption         =   "->"
               Height          =   285
               Left            =   2625
               TabIndex        =   23
               Top             =   0
               Width           =   330
            End
            Begin VB.TextBox TxtCodiceAvvocato 
               Height          =   285
               Left            =   1185
               MaxLength       =   10
               TabIndex        =   22
               Top             =   0
               Width           =   1350
            End
            Begin VB.CommandButton CmdRicercaAnag 
               Caption         =   "&Ricerca Anagrafica"
               Height          =   525
               Left            =   3420
               TabIndex        =   21
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label LblCodAvvocato 
               Caption         =   "Cod. Cassetta:"
               Height          =   255
               Left            =   0
               TabIndex        =   24
               Top             =   30
               Width           =   1110
            End
         End
         Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
            DataField       =   "DataRegistrazione"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoConto.frx":0000
            Caption         =   "StampaEstrattoConto.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoConto.frx":0184
            Keys            =   "StampaEstrattoConto.frx":01A2
            Spin            =   "StampaEstrattoConto.frx":0200
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
            TabIndex        =   15
            Tag             =   "necessario Data Registrazione"
            Top             =   360
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   450
            Calendar        =   "StampaEstrattoConto.frx":0228
            Caption         =   "StampaEstrattoConto.frx":0340
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "StampaEstrattoConto.frx":03AC
            Keys            =   "StampaEstrattoConto.frx":03CA
            Spin            =   "StampaEstrattoConto.frx":0428
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
            TabIndex        =   19
            Top             =   1560
            Width           =   4545
         End
         Begin VB.Label LblDescr 
            Caption         =   "Descrizione:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label LblRicDataFin 
            Caption         =   "Data Fine :"
            Height          =   285
            Left            =   2520
            TabIndex        =   17
            Top             =   120
            Width           =   825
         End
         Begin VB.Label LblRicDataIn 
            Caption         =   "Data Inizio :"
            Height          =   285
            Left            =   135
            TabIndex        =   16
            Top             =   120
            Width           =   870
         End
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Sfratti/Pignor."
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Tag             =   "Pignoramenti"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Notifiche"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         Tag             =   "Notifiche"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Decreti"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Tag             =   "Decreti"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Adempimenti"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Tag             =   "Adempimenti"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ChkAbilitaAnteDef 
         Caption         =   "Abilita anteprima in stampa definitiva"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Frame FrmMetodoStampa 
         Caption         =   "Modalità Stampa"
         Height          =   645
         Left            =   225
         TabIndex        =   4
         Top             =   3465
         Width           =   4920
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
         Height          =   420
         Index           =   0
         Left            =   3375
         TabIndex        =   3
         Top             =   360
         Width           =   1410
      End
      Begin VB.OptionButton OptTipoStampa 
         Caption         =   "Provvisoria"
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   0
         Top             =   450
         Value           =   -1  'True
         Width           =   1320
      End
   End
End
Attribute VB_Name = "StampaEstrattoConto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Avvocato As String
Private WithEvents moFilterManager As CFilterManager
Attribute moFilterManager.VB_VarHelpID = -1
Public TrasferimentoOK As Boolean
Public isUnep As Boolean


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
    Dim avvocatiEstratti As AvvocatiPerEstratto
  Set avvocatiEstratti = GetAvvocatoSingoloPerEstratto(TxtCodiceAvvocato.Text)
   If Not IsDate(TxtRicDataIn.Text) Or Not IsDate(TxtRicDataFin.Text) Then
    MsgBox "Inserire l'intervallo di date", vbOKOnly + vbCritical
    Exit Sub
  End If
  RisolviOrdinamentoErrato
  
  prov = "N"
  If OptTipoStampa(1).value Then prov = "S"
     
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


    
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, _
                             Chk(0), Chk(2), Chk(1), Chk(3), prov, False, 0
           
    
    If Not GetADORecordset("PrtEstrattoConto", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
        If OptTipoStampa(0).value = True Or ChkAbilitaAnteDef.value = True Then
            Dim fb As New FileBackuoHelper
            fb.BackUp g_Settings.AtapUserBackupFolder
            GestStampaDefinitiva
        Else
      
          Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "EstrattoConto.rpt", 1, "Tipo='ESTRATTO'")
          If Stampa.Destination = crptToPrinter Then
                    Unload Stampa
          End If
        End If
    Else
        MsgBox "Nessun dato evaso! Impossibile creare l'Estratto Conto.", vbInformation, "Attenzione"
      
    End If
    DelockPrtTable ("PrtAssegniCircolari")
    DelockPrtTable ("PrtEstrattoConto")
  
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
End Sub
Private Sub moFilterManager_Validate(IsValid As Boolean)
   CmdOK.Enabled = IsValid
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

For i = 0 To 3
  If Chk(i).value = 1 Then schema = schema + Left(Chk(i).Caption, 1)
Next i

nome = TxtCodiceAvvocato.Text
If nome = "" Then nome = "COMPLETO"

TrasferimentoOK = Trasferisci(g_Settings.StoricoECPath & "\EC_" & Format(Date, "yyyymmdd") & "_" & nome & ".mdb", Format(TxtRicDataIn.Text, "yyyymmdd"), Format(TxtRicDataFin.Text, "yyyymmdd"), False, Trim(TxtCodiceAvvocato.Text), schema)
 
Exit Sub

ErroreTrasferimento:
        
    MsgBox "Errore durante trasferimento dati nel db storico!", vbInformation, "Attenzione"
    TrasferimentoOK = False
    Exit Sub
    
End Sub




Public Sub UpdateDataUltimoEstConto()

g_Settings.DBConnection.Execute ("UPDATE Date_EstrattiConto SET DATA_ULTIMO_ESTCONTO='" & Format(TxtRicDataFin.Text, "dd/mm/yyyy") & "'")
    
End Sub
Public Sub AggiornaSaldo(rsAssegni As ADODB.Recordset)
On Error GoTo fine
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
 Set rs = GetADORecordset("Saldi", "chiusura", "codice='" & codice & "'", g_Settings.DBConnection)
 If rs Is Nothing Then
   'Record inesistente
   SQL = "INSERT INTO SALDI(codice,Stato,PROG_Saldi,Commento,SaldoAdemp,SaldoSfpg, " & _
         "SaldoNotif,SaldoDecrIng,SaldoAdempEuro,SaldoSfpgEuro,SaldoNotifEuro,SaldoDecrIngEuro," & _
         "SaldoTotale,SaldoTotaleEuro) " & _
         "VALUES ('" & codice & "','N'," & 1 & ",'" & Commento & "'," & _
         "0,0,0,0,0,0,0,0," & Str(saldo * 1936.27) & "," & Str(saldo) & ");"
 Else
   'record già prersente
   If Format(RitornaData(rs!Chiusura), "yyyy") = Format(dataEC, "yyyy") Then
            
            prog = "PROG_Saldi + 1"
          Else
            prog = 1
            
   End If
   SQL = "UPDATE SALDI SET " & _
         "Stato='N',PROG_Saldi=" & prog & ",Commento='" & Commento & "',SaldoAdemp=0,SaldoSfpg=0, " & _
         "SaldoNotif=0,SaldoDecrIng=0,SaldoAdempEuro=0,SaldoSfpgEuro=0,SaldoNotifEuro=0,SaldoDecrIngEuro=0," & _
         "SaldoTotale=" & Str(saldo * 1936.27) & ",SaldoTotaleEuro=" & Str(saldo) & _
         " WHERE codice='" & codice & "';"
 End If
 g_Settings.DBConnection.Execute SQL
 Exit Sub
fine:
 MsgBox err.Description & vbCrLf & SQL
 
End Sub

Public Sub AggiornaTabellaSaldi()


Dim rsAssegni As ADODB.Recordset


'CreaTabAppAC_Sal  'Crea AssegniCircolari già chiamata prima


Set rsAssegni = GetADORecordset("TempSaldi", "*", "1=1", g_Settings.DBConnection)
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
Public Sub aggiornaFattura(ByRef nFat As Long, codice As String, Data As String, adempi As Double, _
                            decreti As Double, Notifiche As Double, stratti As Double)
Dim SQL As String
Dim rs As ADODB.Recordset
If codice = "525/158" Then
 Debug.Print "Errore"
End If
If GetADORecordset("StoricoFatture", "*", "codAVV='" & codice & "' and DATAFATTURA='" & Format(Data, "yyyymmdd") & "'", g_Settings.DBConnection) Is Nothing Then
     Set rs = GetADORecordset("AnagraficaAvvocati", "*", "codAVV='" & codice & "'", g_Settings.DBConnection)
     
     If rs!AFAT <> "S" Then Exit Sub
     
     SQL = "INSERT INTO StoricoFatture (numOrdinamento,NOME,INDIRI,LOCALI,PROV,CAP,PIVA,codAvv," & _
           "NumeroFattura,DataFattura,DataFatturaNormale,Valuta,ImportoIva,CodIVA, CompAdempEuro,CompDecrIngEuro,CompNotifEuro,CompSfpgEuro) " & _
           "VALUES (" & rs!numOrdinamento & ",'" & Replace(Left(rs!nome, 40), "'", "''") & "','" & Replace(Left(rs!INDIRI, 40), "'", "''") & "','" & Replace(Left(rs!LOCALI, 35), "'", "''") & _
           "','" & rs!prov & "','" & rs!CAP & "','" & rs!PIVA & "','" & codice & "'," & nFat & _
           ",'" & Format(Data, "yyyymmdd") & "','" & Data & "','E'," & g_Settings.IVA & ",'" & g_Settings.CodIVA & "'," & Str(adempi) & "," & Str(decreti) & "," & Str(Notifiche) & "," & Str(stratti) & ");"
           nFat = nFat + 1
   Else
     SQL = "UPDATE StoricoFatture SET " & _
           "CompAdempEuro=CompAdempEuro+" & Str(adempi) & _
           ",CompDecrIngEuro=CompDecrIngEuro+" & Str(decreti) & _
           ",CompNotifEuro=CompNotifEuro+" & Str(Notifiche) & _
           ",CompSfpgEuro=CompSfpgEuro+" & Str(stratti) & _
           " WHERE codAVV='" & codice & "' and DATAFATTURA='" & Data & "';"
   
End If
g_Settings.DBConnection.Execute SQL

End Sub
Public Sub GeneraFattura(Numero As Integer, Data As Date)
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


SQL = "SELECT codAvv,DESCR_ATTIVITA,Sum(Competenze) FROM PrtEstrattoConto " & _
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
     aggiornaFattura nFat, codice, "" & Data, adempi, decreti, Notifiche, sfratti
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
     aggiornaFattura nFat, codice, "" & Data, adempi, decreti, Notifiche, sfratti
End If
End Sub
Private Sub SalvaSaldiTemporanei()

Dim qry As String

Dim sp1, sp3, sp5 As Double
On Error GoTo fine


'Reset PrtEstrattoConto
qry = "DELETE * FROM TempSaldi;"
g_Settings.DBConnection.Execute qry

qry = GetQuerySaldi("TempSaldi", " < ")
g_Settings.DBConnection.Execute qry

qry = GetQuerySaldi("TempSaldi", " >= ")
g_Settings.DBConnection.Execute qry

Exit Sub
fine:
 MsgBox err.Description

End Sub
Private Sub CreaTabAssegni()

Dim qry As String

Dim sp1, sp3, sp5 As Double
On Error GoTo fine


'Reset PrtEstrattoConto
qry = "DELETE * FROM PrtAssegniCircolari;"
g_Settings.DBConnection.Execute qry
qry = GetQuerySaldi("PrtAssegniCircolari", " >= ")
g_Settings.DBConnection.Execute qry
Exit Sub
fine:
 MsgBox err.Description

End Sub
Private Function GetQuerySaldi(destinationTable As String, condition As String) As String
Dim qry As String

qry = "INSERT INTO " & destinationTable & " ( CODAVV, NOME, DESCR_ATTIVITA, DEPOSITO, COMPETENZE, SALDO, " & _
     "SPESE1, SPESE2, SPESE3, SPESE4, SPESE5, SPESE6, SALDO_PRECEDENTE, VALUTA,NumOrdinamento,DATA_INIZIO,DATA_FINE ) " & _
     "SELECT CODAVV, NOME, 'XXX', Sum(PrtEstrattoConto.DEPOSITO) AS DEP," & _
     "Sum(PrtEstrattoConto.COMPETENZE)*" & Str(1 + g_Settings.IVA / 100) & " AS [COMP], [DEP]-[COMP]-[S1]-[S2]-[S3]-[S4]-[S5]-[S6] AS Ass," & _
     "Sum(IIF(DESCR_SPESE1='Fotocopie',[SPESE1]*[PrtEstrattoConto]![QtaFotocopie],[SPESE1])) AS S1, Sum(PrtEstrattoConto.SPESE2) AS S2," & _
     "Sum(IIF(DESCR_SPESE3='Marche',[SPESE3]*[QtaMarche],[SPESE3])) AS S3, Sum(PrtEstrattoConto.SPESE4) AS S4, " & _
     "Sum(IIF(DESCR_SPESE5='Diritti Cancelleria',[SPESE5]*[qtaDirittiCancelleria],[SPESE5])) AS S5, Sum(PrtEstrattoConto.SPESE6) AS S6," & _
     "fIRST(PrtEstrattoConto.SALDO_PRECEDENTE) AS S_PRECEDENTE, 'E' AS Valuta,NumOrdinamento,DATA_INIZIO,DATA_FINE " & _
     "From PrtEstrattoConto " & _
     "GROUP BY PrtEstrattoConto.CODAVV,PrtEstrattoConto.Saldo_Precedente, PrtEstrattoConto.NOME, NumOrdinamento,DATA_INIZIO,DATA_FINE " & _
     " HAVING   Sum(PrtEstrattoConto.DEPOSITO) + fIRST(PrtEstrattoConto.SALDO_PRECEDENTE) -Sum(PrtEstrattoConto.COMPETENZE)*" & Str(1 + g_Settings.IVA / 100) & "-" & _
     "Sum(IIF(DESCR_SPESE1='Fotocopie',[SPESE1]*[PrtEstrattoConto]![QtaFotocopie],[SPESE1]))-" & _
     "Sum(PrtEstrattoConto.SPESE2) -Sum(IIF(DESCR_SPESE3='Marche',[SPESE3]*[QtaMarche],[SPESE3])) - " & _
     "Sum(PrtEstrattoConto.SPESE4) - Sum(IIF(DESCR_SPESE5='Diritti Cancelleria',[SPESE5]*[qtaDirittiCancelleria],[SPESE5])) - " & _
     "Sum(PrtEstrattoConto.SPESE6) " & condition & Str(g_Settings.LimiteSaldo)
  GetQuerySaldi = qry
End Function

Public Sub CreazioneStampaAssegniCircolari()


SalvaSaldiTemporanei

CreaTabAssegni


g_Settings.DBConnection.Execute "DELETE * FROM PrtAssegniCircolari where (saldo + SALDO_PRECEDENTE)<" & Str(g_Settings.LimiteSaldo)

End Sub

Public Sub GestStampaDefinitiva()

Dim MSG_Avviso, Response As Variant
    MSG_Avviso = "Durante questa operazione è necessario non modificare alcun dato." & Chr(10)
    MSG_Avviso = MSG_Avviso & "Chiudere tutte le finestre aperte e verificare che nessun"
    MSG_Avviso = MSG_Avviso & " altro client abbia l'applicazione attiva!" & Chr(10) & "Proseguire?"
    Response = MsgBox(MSG_Avviso, vbYesNo + vbInformation + vbDefaultButton1, "Avviso")
    If Response = vbYes Then    ' User chose Yes.
            Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "EstrattoConto.rpt", 2, "Tipo='ESTRATTO'")
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
                Call Stampa.gestioneReport("", "", 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "AssegniCircolari.rpt", 3)
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
                        ImpostazioniFatturazione.isUnep = False
                        ImpostazioniFatturazione.Show
                    End If
               End If
           End If
   End If
End Sub

