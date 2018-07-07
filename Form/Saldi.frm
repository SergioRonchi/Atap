VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form Saldi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Saldi"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11700
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   450
      Left            =   10440
      TabIndex        =   25
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton CmdElimina 
      Caption         =   "&Elimina"
      Enabled         =   0   'False
      Height          =   450
      Left            =   9120
      TabIndex        =   24
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame FrmRicercaSaldi 
      Height          =   5265
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   11595
      Begin VB.TextBox TxtRicerca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   3
         Top             =   120
         Width           =   3720
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "&Ricerca"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid flex 
         Height          =   4575
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   11295
         _cx             =   19923
         _cy             =   8070
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame FrmSaldi 
      Height          =   2280
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      Begin VB.CommandButton cmdElimina2 
         Caption         =   "&Elimina"
         Height          =   570
         Left            =   10200
         TabIndex        =   28
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "Stampa"
         Height          =   375
         Left            =   10200
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Nuovo"
         Height          =   375
         Left            =   10200
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   570
         Left            =   10200
         TabIndex        =   23
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox TxtCodice 
         DataField       =   "CODAVV"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   315
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate txtDataChiusura 
         DataField       =   "Chiusura"
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calendar        =   "Saldi.frx":0000
         Caption         =   "Saldi.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":0184
         Keys            =   "Saldi.frx":01A2
         Spin            =   "Saldi.frx":0200
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
      Begin TDBNumber6Ctl.TDBNumber txtSaldo 
         DataField       =   "SaldoAdempEuro"
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Saldi.frx":0228
         Caption         =   "Saldi.frx":0248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":02B4
         Keys            =   "Saldi.frx":02D2
         Spin            =   "Saldi.frx":031C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtSaldo 
         DataField       =   "SaldoNotifEuro"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Saldi.frx":0344
         Caption         =   "Saldi.frx":0364
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":03D0
         Keys            =   "Saldi.frx":03EE
         Spin            =   "Saldi.frx":0438
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtSaldo 
         DataField       =   "SaldoTotaleEuro"
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Saldi.frx":0460
         Caption         =   "Saldi.frx":0480
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":04EC
         Keys            =   "Saldi.frx":050A
         Spin            =   "Saldi.frx":0554
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999.99
         MinValue        =   -999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtSaldo 
         DataField       =   "SaldoSfpgEuro"
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Saldi.frx":057C
         Caption         =   "Saldi.frx":059C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":0608
         Keys            =   "Saldi.frx":0626
         Spin            =   "Saldi.frx":0670
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtSaldo 
         DataField       =   "SaldoDecrIngEuro"
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   21
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         Calculator      =   "Saldi.frx":0698
         Caption         =   "Saldi.frx":06B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Saldi.frx":0724
         Keys            =   "Saldi.frx":0742
         Spin            =   "Saldi.frx":078C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "#,##0.00;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,##0.00"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label SaldiNegativi 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label LblPrezzo5 
         Caption         =   "Saldo Totale"
         Height          =   285
         Left            =   450
         TabIndex        =   14
         Top             =   1800
         Width           =   1635
      End
      Begin VB.Label LblDescrizioneAvvocato 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   705
         Width           =   7410
      End
      Begin VB.Label LblData 
         Caption         =   "Data chiusura mensile"
         Height          =   285
         Left            =   3975
         TabIndex        =   12
         Top             =   345
         Width           =   1770
      End
      Begin VB.Label LblCodice 
         Caption         =   "Cod. Cassetta :"
         Height          =   285
         Left            =   450
         TabIndex        =   10
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label LblDescrizione 
         Caption         =   "Descrizione"
         Height          =   285
         Left            =   450
         TabIndex        =   9
         Top             =   720
         Width           =   960
      End
      Begin VB.Label LblPrezzo1 
         Caption         =   "Saldo Adempimenti"
         Height          =   330
         Left            =   480
         TabIndex        =   8
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label LblPrezzo2 
         Caption         =   "Saldo Sfratti/Pignoramenti"
         Height          =   285
         Left            =   3825
         TabIndex        =   7
         Top             =   1140
         Width           =   1950
      End
      Begin VB.Label LblPrezzo3 
         Caption         =   "Saldo Notifiche"
         Height          =   285
         Left            =   450
         TabIndex        =   6
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label LblValutaCor 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   1800
         Width           =   1980
      End
      Begin VB.Label LblPrezzo4 
         Caption         =   "Saldo Decreti Ingiuntivi"
         Height          =   285
         Left            =   3825
         TabIndex        =   4
         Top             =   1590
         Width           =   1860
      End
   End
End
Attribute VB_Name = "Saldi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const EURO = 1936.27
Dim qry As String
Dim qryOrder As String
Dim qryWhere As String
Public Tabella As String
Public Campo1 As String
Public Campo2 As String

Public isUnep As Boolean
Dim TabellaSaldi As String
Dim Data As String



Private Sub cmdElimina_Click()
Dim Response As Long
Response = MsgBox("Vuoi eliminare il record  " & TxtCodice & " ?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
If Response = vbYes Then    ' User chose Yes.
g_Settings.DBConnection.Execute "DELETE * FROM " & TabellaSaldi & " WHERE Codice='" & TxtCodice & "'"
MsgBox "Record " & TxtCodice & " eliminato!"
Aggiorna
End If
End Sub


Private Sub cmdPrint_Click()
  PrintForm
End Sub

Private Sub cmdElimina2_Click()
Dim Response As Long
Dim RecCor As Long
Dim c As Control
Dim SQL As String
    RecCor = flex.row
    'Sto Modificando la mia anagrafica
    Response = MsgBox("Vuoi eliminare il saldo selezionato?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
     SQL = ""
         
         SQL = "DELETE * FROM " & TabellaSaldi
         SQL = SQL & " WHERE codice='" & TxtCodice & "'"
         g_Settings.DBConnection.Execute SQL
         Aggiorna
         On Error GoTo prox
         flex.row = RecCor
prox:
         flex_DblClick
    End If

End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
PulisciCampi Me
LblDescrizioneAvvocato = ""
End Sub

Private Sub CmdRicerca_Click()
  
    If TxtRicerca.Text <> "" Then
        qryWhere = " WHERE (codice  Like '" & TxtRicerca.Text & "%' OR AnagraficaAvvocati.NOME Like '" & TxtRicerca.Text & "%') "
       Else
        qryWhere = ""
    End If
   Aggiorna
End Sub

Private Sub CmdSalva_Click()
Dim Response As Long
Dim RecCor As Long
Dim c As Control
Dim SQL As String
    RecCor = flex.row
    'Sto Modificando la mia anagrafica
    Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
     SQL = ""
         For Each c In Me.Controls
          If TypeOf c Is TDBNumber Then
            If c.DataField <> "" Then SQL = SQL & c.DataField & "='" & c.Text & "',"
          End If
         Next
         SQL = "UPDATE " & TabellaSaldi & " SET " & SQL
         SQL = SQL & " Chiusura='" & ElaboraData(txtDataChiusura) & "'"
         If txtSaldo(2).value < -g_Settings.LimiteSaldo Then
           SQL = SQL & ",Commento='Saldo Negativo'"
          Else
           SQL = SQL & ",Commento=' '"
         End If
         SQL = SQL & " WHERE codice='" & TxtCodice & "'"
         g_Settings.DBConnection.Execute SQL
         Aggiorna
         flex.row = RecCor
         flex_DblClick
    End If
End Sub



Private Sub cmdStampa_Click()
PrintForm
End Sub

Private Sub flex_BeforeSort(ByVal Col As Long, Order As Integer)
    
    Call sortGrid(flex, Col, Order, 1, -1)


End Sub

Private Sub flex_DblClick()
Dim SQL As String
Dim txt As TDBNumber
Dim rs As ADODB.Recordset
Dim D As Double
Dim r As Long
r = flex.row
If r = 0 Then Exit Sub
TxtCodice = flex.TextMatrix(r, 1)
LblDescrizioneAvvocato = flex.TextMatrix(r, 2)
Set rs = GetADORecordset(TabellaSaldi, "*", "Codice='" & TxtCodice & "' ", g_Settings.DBConnection)
                               
For Each txt In Me.txtSaldo
If (IsNumeric(rs(txt.DataField))) Then
   D = Round(rs(txt.DataField), 2)
  
  txt.value = D
End If
Next
txtDataChiusura.Text = RitornaData(rs!Chiusura)
SaldiNegativi.Caption = ControlloNULL(rs!Commento)
End Sub

Private Sub Form_Load()

    If isUnep Then
     Caption = "Saldi UNEP"
     TabellaSaldi = "SaldiUNEP"
    Else
      Caption = "Saldi"
      TabellaSaldi = "Saldi"
    End If
    qry = " SELECT " & TabellaSaldi & ".Codice, AnagraficaAvvocati.NOME, " & _
          "Mid(" & TabellaSaldi & ".Chiusura,7,2) & '/' & Mid(" & TabellaSaldi & ".Chiusura,5,2) & '/' & Mid(" & TabellaSaldi & ".Chiusura,1,4) as Data, " & TabellaSaldi & ".Chiusura, " & _
          "" & TabellaSaldi & ".SaldoAdempEuro as Adempi, " & TabellaSaldi & ".SaldoSfpgEuro as Sfratti, " & TabellaSaldi & ".SaldoNotifEuro as Notifiche, " & TabellaSaldi & ".SaldoDecrIngEuro as Decreti, " & TabellaSaldi & ".SaldoTotaleEuro As Totale, Commento,AnagraficaAvvocati.NumOrdinamento "
    qry = qry & " FROM " & TabellaSaldi & " INNER JOIN AnagraficaAvvocati ON " & TabellaSaldi & ".Codice = AnagraficaAvvocati.CODAVV "
          
    qryOrder = " ORDER BY " & TabellaSaldi & ".numOrdinamento"
    Aggiorna
End Sub
Public Sub Aggiorna()
 AggiornaGriglia flex, qry & qryWhere & qryOrder, cmdElimina
 flex.ColWidth(2) = 3000
 flex.ColWidth(9) = 1000
 flex.ColWidth(10) = 1800
 Dim I As Integer
 For I = 5 To 9
   flex.ColFormat(I) = "#,##0.00"
 Next I
 flex.ColHidden(4) = True
 flex.ColHidden(11) = True
 
 flex.ColDataType(3) = flexDTDate
 flex.ColSort(3) = flexSortUseColSort
 
 flex_DblClick
End Sub


Public Sub CalcolaSaldoTotale()
Dim I As Integer, p As Integer

Dim saldo As Double
    saldo = 0
    For I = 0 To txtSaldo.count - 1
  If txtSaldo(I).DataField <> "SaldoTotaleEuro" Then
    saldo = saldo + txtSaldo(I).value
   Else
    p = I
  End If
Next I
    
    txtSaldo(p).value = saldo

End Sub

Private Sub txtSaldo_Change(index As Integer)
  If index <> 2 Then CalcolaSaldoTotale
End Sub


Private Sub txtSaldo_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub
