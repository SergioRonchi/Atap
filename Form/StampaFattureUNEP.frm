VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form StampaFattureUNEP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Fatture UNEP"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictureUNEP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "StampaFattureUNEP.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   14
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "E&sci"
      Height          =   500
      Left            =   4560
      TabIndex        =   7
      Top             =   8160
      Width           =   1380
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   500
      Left            =   3120
      TabIndex        =   6
      Top             =   8160
      Width           =   1380
   End
   Begin VB.Frame FrmRicercaTribunali 
      Height          =   7245
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   5910
      Begin VB.CheckBox Check1 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TxtRicercaData 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "Tutti"
         Top             =   200
         Width           =   1395
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "&Ricerca"
         Height          =   375
         Left            =   4440
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid flex 
         Height          =   6015
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5415
         _cx             =   9551
         _cy             =   10610
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
         SelectionMode   =   3
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
      Begin TDBDate6Ctl.TDBDate txtDataDa 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "necessario Data Registrazione"
         Top             =   600
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         Calendar        =   "StampaFattureUNEP.frx":0442
         Caption         =   "StampaFattureUNEP.frx":055A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaFattureUNEP.frx":05BA
         Keys            =   "StampaFattureUNEP.frx":05D8
         Spin            =   "StampaFattureUNEP.frx":0636
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
         MaxDate         =   44196
         MinDate         =   36161
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
      Begin TDBDate6Ctl.TDBDate txtDataA 
         DataField       =   "DataRegistrazione"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Tag             =   "necessario Data Registrazione"
         Top             =   600
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         Calendar        =   "StampaFattureUNEP.frx":065E
         Caption         =   "StampaFattureUNEP.frx":0776
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "StampaFattureUNEP.frx":07D6
         Keys            =   "StampaFattureUNEP.frx":07F4
         Spin            =   "StampaFattureUNEP.frx":0852
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
         Enabled         =   0
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
         MaxDate         =   44196
         MinDate         =   36161
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
      Begin VB.Label Label1 
         Caption         =   "Codice"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrmMetodoStampa 
      Caption         =   "Modalità Stampa"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   7320
      Width           =   5670
      Begin VB.CheckBox chkTutti 
         Caption         =   "Stampa tutte"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton OptModSt 
         Caption         =   "Anteprima"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton OptModSt 
         Caption         =   "Diretta"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "StampaFattureUNEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qrySQL As String

Private Sub Check1_Click()
 txtDataA.Enabled = (Check1.value = 1)
End Sub

Private Sub CmdAnnulla_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub

Private Sub CmdOK_Click()
    createSelectionFormula
    Call Stampa.gestioneReport("", qrySQL, 0, IIf(OptModSt(0).value, crptToWindow, crptToPrinter), "FatturaUNEP.rpt", 2)
End Sub


Private Sub createSelectionFormula()
 Dim i As Integer
 Dim r As Integer
 Dim key As String
 
 qrySQL = " {StoricoFatture.DataFattura} in """ & Format(txtDataDa.Text, "yyyymmdd") & """ to """ & Format(txtDataA.Text, "yyyymmdd") & """"
 
  If chkTutti.value <> 1 Then
     qrySQL = " 1=1 "
     If flex.SelectedRows > 0 Then
            
            qrySQL = qrySQL & " and ("
            For i = 0 To flex.SelectedRows - 2
              r = flex.SelectedRow(i)
              
              qrySQL = qrySQL & "({StoricoFatture.CODAVV}=""" & flex.TextMatrix(r, 1) & """ and " & _
                              " {StoricoFatture.NumeroFattura}=""" & flex.TextMatrix(r, 5) & """ and " & _
                              " {StoricoFatture.DataFattura}=""" & flex.TextMatrix(r, 6) & """) or "
            Next i
            r = flex.SelectedRow(i)
            qrySQL = qrySQL & "({StoricoFatture.CODAVV}=""" & flex.TextMatrix(r, 1) & """ and " & _
                              " {StoricoFatture.NumeroFattura}=""" & flex.TextMatrix(r, 5) & """ and " & _
                              " {StoricoFatture.DataFattura}=""" & flex.TextMatrix(r, 6) & """) )"
    End If
  End If
End Sub



Private Sub CmdRicerca_Click()
    Dim qry As String
    Dim qry1 As String
     
    qry = "SELECT CODAVV,DataFatturaNormale As Data,Nome as Avvocato,NumOrdinamento,NumeroFattura,datafattura FROM StoricoFattureUNEP "
    If TxtRicercaData.Text <> "" Then
        If IsDate(txtDataDa) Then
           qry1 = " DataFattura >= '" & Format(txtDataDa.Text, "yyyymmdd") & "'"
        End If
        If IsDate(txtDataA) Then
         qry1 = qry1 & " And DataFattura <= '" & Format(txtDataA.Text, "yyyymmdd") & "'"
        End If
        
        If TxtRicercaData <> "Tutti" And TxtRicercaData <> "" Then
          qry1 = qry1 & " AND CodAvv Like '" & TxtRicercaData & "%' "
        End If
        qry = qry + " WHERE " + qry1
    End If
    qry = qry + " Order By DataFatturaNormale DESC, NumOrdinamento; "
     
    AggiornaGriglia flex, qry
    flex.ColWidth(1) = 1200
    flex.ColWidth(2) = 1100
    flex.ColWidth(3) = 2550
    flex.ColHidden(4) = True
    flex.ColHidden(5) = True
    flex.ColHidden(6) = True
    flex.ColDataType(2) = flexDTDate
End Sub

Private Sub flex_BeforeSort(ByVal Col As Long, Order As Integer)
Call sortGrid(flex, Col, Order, 1, 4)
End Sub

Private Sub Form_Load()
    
    getDataFattura
    CmdRicerca_Click
End Sub



Private Sub getDataFattura()
    txtDataDa.value = GetADOValue("Date_EstrattiConto", "DATA_FATTURA_UNEP", "1=1", g_Settings.DBConnection)
    
    txtDataA.value = txtDataDa.value
End Sub

Private Sub txtDataDa_Change()
 If txtDataA.Enabled = False And IsDate(txtDataDa) Then
   txtDataA = txtDataDa
 End If
End Sub
