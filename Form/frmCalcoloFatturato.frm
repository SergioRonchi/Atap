VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmCalcoloFatturato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calcolo Fatturato"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optUnep 
      Caption         =   "UNEP"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton optAtap 
      Caption         =   "Atap"
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo fatturato UNEP"
      Height          =   615
      Left            =   3960
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton optMese 
         Caption         =   "Mese"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMese 
         Caption         =   "Bimestre"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "Esci"
      Height          =   450
      Left            =   5400
      TabIndex        =   20
      Top             =   5880
      Width           =   1860
   End
   Begin VB.CommandButton cmdCalcola 
      Caption         =   "Calcola"
      Height          =   525
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate TxtRicDataIn 
      DataField       =   "DataRegistrazione"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "necessario Data Registrazione"
      Top             =   360
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calendar        =   "frmCalcoloFatturato.frx":0000
      Caption         =   "frmCalcoloFatturato.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCalcoloFatturato.frx":0184
      Keys            =   "frmCalcoloFatturato.frx":01A2
      Spin            =   "frmCalcoloFatturato.frx":0200
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
      TabIndex        =   1
      Tag             =   "necessario Data Registrazione"
      Top             =   360
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calendar        =   "frmCalcoloFatturato.frx":0228
      Caption         =   "frmCalcoloFatturato.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCalcoloFatturato.frx":03AC
      Keys            =   "frmCalcoloFatturato.frx":03CA
      Spin            =   "frmCalcoloFatturato.frx":0428
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
   Begin VB.Label lblQuote 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   27
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblQuote 
      Alignment       =   1  'Right Justify
      Caption         =   "Quote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   26
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Lordo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   18
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Imponibile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   17
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblFatAdempi 
      Alignment       =   1  'Right Justify
      Caption         =   "Adempimenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblFatNotifiche 
      Alignment       =   1  'Right Justify
      Caption         =   "Notifiche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblFatSfratti 
      Alignment       =   1  'Right Justify
      Caption         =   "Sfratti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblFatDecreti 
      Alignment       =   1  'Right Justify
      Caption         =   "Decreti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   13
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblFatLordo 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblFatIVA 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   11
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblImpo 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblFatAdempi 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblFatNotifiche 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblFatSfratti 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblFatDecreti 
      Alignment       =   1  'Right Justify
      Caption         =   "€ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fatturato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label LblRicDataIn 
      Caption         =   "Data Inizio :"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
   Begin VB.Label LblRicDataFin 
      Caption         =   "Data Fine :"
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmCalcoloFatturato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnulla_Click()
Unload Me
End Sub

Private Sub cmdCalcola_Click()
     If IsPrtTableLocked("PrtAssegniCircolari") Or IsPrtTableLocked("PrtEstrattoConto") Or _
        IsPrtTableLocked("PrtAssegniCircolariUNEP") Or IsPrtTableLocked("PrtEstrattoContoUNEP") Then
      MsgBox "Attenzione: " & vbCrLf & _
             "E' già in corso una stampa che riguarda i dati selezionati." & vbCrLf & _
             "Si prega di riprovare tra qualche istante." & vbCrLf & vbCrLf & _
             "Se il problema persiste e non sono in corso altre stampe si consiglia di:" & vbCrLf & _
             " - Eseguire 'Sblocca Stampe' dal menu 'Utilità'", vbInformation + vbOKOnly
      Exit Sub
    End If
  Dim avvocatiEstratti As AvvocatiPerEstratto
  If Not IsDate(TxtRicDataIn.Text) Or Not IsDate(TxtRicDataFin.Text) Then
    MsgBox "Inserire l'intervallo di date", vbOKOnly + vbCritical
    Exit Sub
  End If
  Dim sql As String
  
  Set avvocatiEstratti = GetAvvocatoSingoloPerEstratto("")
  If optAtap Then
    CalcolaNormale avvocatiEstratti
        sql = "SELECT Sum(CompAdempEuro) AS Ademp, " & _
          "Sum(CompSfpgEuro) AS Sfratti, " & _
          "Sum(CompNotifEuro) AS Notifiche, " & _
          "Sum(CompNotifEuro) AS Decreti, " & _
          "Sum(0) AS Quota, " & _
          "Sum(CompAdempEuro+CompSfpgEuro+CompNotifEuro+CompNotifEuro) as TotaleImponibile, " & _
          "Sum((CompAdempEuro+CompSfpgEuro+CompNotifEuro+CompNotifEuro)*ImportoIVA/100)as TotaleIVA, " & _
          "Sum((CompAdempEuro + CompSfpgEuro + CompNotifEuro + CompNotifEuro) * (1 + ImportoIVA / 100)) As TotaleLordo " & _
          "FROM FattureTemp;"
   Else
    CalcolaUnep avvocatiEstratti
          sql = "SELECT Sum(CompAdempEuro) AS Ademp, " & _
          "Sum(CompSfpgEuro) AS Sfratti, " & _
          "Sum(CompNotifEuro) AS Notifiche, " & _
          "Sum(CompNotifEuro) AS Decreti, " & _
          "Sum(Quota) AS Quota, " & _
          "Sum(CompAdempEuro+CompSfpgEuro+CompNotifEuro+CompNotifEuro+Quota) as TotaleImponibile, " & _
          "Sum((CompAdempEuro+CompSfpgEuro+CompNotifEuro+CompNotifEuro+Quota)*ImportoIVA/100)as TotaleIVA, " & _
          "Sum((CompAdempEuro + CompSfpgEuro + CompNotifEuro + CompNotifEuro+Quota) * (1 + ImportoIVA / 100)) As TotaleLordo " & _
          "FROM FattureTempUNEP;"
  End If
  
   
   
    Dim rs As ADODB.Recordset
   Set rs = newAdoRs

        rs.Open sql, g_Settings.DBConnection
    
        If Not rs.EOF Then
         PopolaLabel 0, rs
        End If
        rs.Close
   

End Sub
Private Sub CalcolaNormale(avvocatiEstratti As AvvocatiPerEstratto)

    LockPrtTable ("PrtAssegniCircolari")
    LockPrtTable ("PrtEstrattoConto")
     OpenProgress ("Attendere... Preparazione Fatturato!")
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, 1, 1, 1, 1, "N", False, 0, ""
    
    If Not GetADORecordset("PrtEstrattoConto", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
        
       StampaEstrattoConto.CreazioneStampaAssegniCircolari
       StampaEstrattoConto.GeneraFattura "0", Format(Now, "DD/MM/YYYY"), True
     
    End If
   
       DelockPrtTable ("PrtAssegniCircolari")
    DelockPrtTable ("PrtEstrattoConto")
End Sub
Private Sub CalcolaUnep(avvocatiEstratti As AvvocatiPerEstratto)

    LockPrtTable ("PrtAssegniCircolariUNEP")
    LockPrtTable ("PrtEstrattoContoUNEP")
     OpenProgress ("Attendere... Preparazione Fatturato UNEP!")
     
     
    g_Settings.DBConnection.Execute "DELETE * FROM PrtData"
    
    g_Settings.DBConnection.Execute "INSERT INTO PrtData(Tipo, Bimestre, BimestreAnno) VALUES(" & IIf(optMese(0).value, 1, 2) & ",1,2)"
    
    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, 0, 1, 0, 1, "N", True, IIf(optMese(0).value, 1, 2), ""
   
    AggiungiDeduzioni TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti
    
    AggiungiAvvocatiQuota TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, IIf(optMese(0).value, g_Settings.QuotaSoci / 2, g_Settings.QuotaSoci)

    Riempi_PRT_EstrattoContoX TxtRicDataIn.Text, TxtRicDataFin.Text, avvocatiEstratti, 1, 1, 1, 1, "N", False, 0, ""
    
    If Not GetADORecordset("PrtEstrattoContoUNEP", "*", "1=1", g_Settings.DBConnection) Is Nothing Then
        
       StampaEstrattoContoUNEP.CreazioneStampaAssegniCircolari
       StampaEstrattoContoUNEP.GeneraFattura "0", Format(Now, "DD/MM/YYYY"), True
     
    End If
   
       DelockPrtTable ("PrtAssegniCircolari")
    DelockPrtTable ("PrtEstrattoConto")
End Sub

Private Sub PopolaLabel(index As Integer, rs As ADODB.Recordset)

  lblQuote(index).Caption = Format(rs("Quota"), "#,##0.00")
  lblFatAdempi(index).Caption = Format(rs("Ademp"), "#,##0.00")
  lblFatDecreti(index).Caption = Format(rs("Decreti"), "#,##0.00")
  lblFatNotifiche(index).Caption = Format(rs("Notifiche"), "#,##0.00")
  lblFatSfratti(index).Caption = Format(rs("Sfratti"), "#,##0.00")
  lblImpo(index).Caption = Format(rs("TotaleImponibile"), "#,##0.00")
  lblFatIVA(index).Caption = Format(rs("TotaleIVA"), "#,##0.00")
  lblFatLordo(index).Caption = Format(rs("TotaleLordo"), "#,##0.00")
  
End Sub
Private Sub Form_Load()
 TxtRicDataIn = #1/1/1999#
End Sub

Private Sub optAtap_Click()
Frame2.Visible = optUnep.value = True
lblQuote(3).Visible = optUnep.value = True
lblQuote(0).Visible = optUnep.value = True

lblFatDecreti(3).Visible = optUnep.value = False
lblFatDecreti(0).Visible = optUnep.value = False
lblFatAdempi(3).Visible = optUnep.value = False
lblFatAdempi(0).Visible = optUnep.value = False
End Sub

Private Sub optUnep_Click()
 Frame2.Visible = optUnep.value = True
 lblQuote(3).Visible = optUnep.value = True
 lblQuote(0).Visible = optUnep.value = True
 
lblFatDecreti(3).Visible = optUnep.value = False
lblFatDecreti(0).Visible = optUnep.value = False
lblFatAdempi(3).Visible = optUnep.value = False
lblFatAdempi(0).Visible = optUnep.value = False
End Sub
