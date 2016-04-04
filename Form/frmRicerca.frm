VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmRicerca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ricerca"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13080
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   495
      Left            =   11760
      TabIndex        =   7
      Top             =   9720
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid flex 
      Height          =   8055
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   12975
      _cx             =   22886
      _cy             =   14208
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
      AllowUserResizing=   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraAna 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   12795
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Nuova Cassetta"
         Height          =   375
         Left            =   11280
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdFiltraAna 
         Caption         =   "Filtra"
         Height          =   615
         Left            =   11280
         Picture         =   "frmRicerca.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxtUsufruente 
         Height          =   330
         Left            =   1785
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1005
         Width           =   3870
      End
      Begin VB.TextBox TxtRicNome 
         Height          =   330
         Left            =   1785
         MaxLength       =   40
         TabIndex        =   13
         Top             =   525
         Width           =   3870
      End
      Begin VB.TextBox TxtRicCodAvvInt 
         Height          =   330
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label LblUsufruente 
         Caption         =   "Usufruente :"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label LblRicNome 
         Caption         =   "Cognome e Nome :"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label LblRicCodAvvInt 
         Caption         =   "Cod. Cassetta :"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   165
         Width           =   1500
      End
   End
   Begin VB.Frame fraTop 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   960
         Width           =   3495
      End
      Begin VB.ComboBox cmbSiglaCh 
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   550
         Width           =   1725
      End
      Begin VB.ComboBox cmbSigla 
         Height          =   315
         Left            =   6000
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   170
         Width           =   1725
      End
      Begin TDBDate6Ctl.TDBDate Da 
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   960
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Calendar        =   "frmRicerca.frx":014A
         Caption         =   "frmRicerca.frx":0262
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRicerca.frx":02CE
         Keys            =   "frmRicerca.frx":02EC
         Spin            =   "frmRicerca.frx":034A
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
      Begin VB.OptionButton opt 
         Caption         =   "Inevasi"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Evasi"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Tutto"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   8
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdFiltra 
         Caption         =   "Filtra"
         Height          =   735
         Left            =   11280
         Picture         =   "frmRicerca.frx":0372
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxtRicCodAvv 
         Height          =   285
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
      Begin TDBDate6Ctl.TDBDate A 
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   960
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Calendar        =   "frmRicerca.frx":04BC
         Caption         =   "frmRicerca.frx":05D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRicerca.frx":0640
         Keys            =   "frmRicerca.frx":065E
         Spin            =   "frmRicerca.frx":06BC
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
      Begin VB.Label Label2 
         Caption         =   "Chiusura"
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Inserimento"
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblRicDataFin 
         Caption         =   "Data Fine :"
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label LblRicDataIn 
         Caption         =   "Data Inizio :"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label LblRicCodAvv 
         Caption         =   "Cod. Cassetta :"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   405
         Width           =   1410
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuElimina 
         Caption         =   "Elimina"
      End
   End
End
Attribute VB_Name = "frmRicerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmCaller As IForm
Public Query As String
Public Tabella As String
Public Titolo As String
Public DefaultOrder As String
Public NCol As Integer
Public PosizioneCodice As Integer
Public tipo As String
Public Filtro As String
Public Location As Long
Dim qryExe As String
Private mFoundCode As String
Public isUnep As Boolean
Public Event AvvocatoSelezionato(codice As String)

Public Property Get FoundCode() As String
  FoundCode = mFoundCode
End Property





Private Sub cmbDate_Click()
  Dim data1 As Date
  Dim data2 As Date
  data2 = Date + 30
  data2 = LastDay(month(data2), year(data2))
  A = data2
 
Select Case cmbDate.ListIndex
  Case 0 'mese
    data1 = Date - 30
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 1 'trimestre
    data1 = Date - 90
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 2 'Semestre
    data1 = Date - 180
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 3 'Anno
    data1 = Date - 365
    Da = 1 & "/" & month(data1) & "/" & year(data1)
  Case 4 'Anno Completo
    Da = 1 & "/" & 1 & "/" & year(Date)
    A = 31 & "/" & 12 & "/" & year(Date)
  Case 5 'Tutto
    Da = ""
    A = ""
End Select


cmdFiltra_Click
End Sub

Private Sub cmbSigla_Click()
 cmdFiltra_Click
End Sub


Private Sub cmbSigla_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub cmbSiglaCh_Click()
cmdFiltra_Click
End Sub

Private Sub cmbSiglaCh_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
 AnagAvvocati.Azione = TipoAzione.Nuovo
 AnagAvvocati.Show
End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdFiltra_Click()
qryExe = Query
qryExe = qryExe & " WHERE 1=1 " & Filtro
    If TxtRicCodAvv.Text <> "" Then
        qryExe = qryExe & " AND (CODAVV = '" & TxtRicCodAvv.Text & "')"
     End If
    If IsDate(Da) Then
        qryExe = qryExe & " AND ( DataRegistrazione >= '" & Format(Da, "yyyymmdd") & "')"
    End If
    If IsDate(A) Then
        qryExe = qryExe & " AND ( DataRegistrazione <= '" & Format(A, "yyyymmdd") & "')"
    End If
    If opt(1) Then
       qryExe = qryExe & " AND ( CheckVisual = 'X')"
    End If
    If opt(2) Then
       qryExe = qryExe & " AND ( CheckVisual <> 'X')"
    End If
    If cmbSigla.ListIndex > 0 Then
       qryExe = qryExe & " AND ( SIGLA ='" & cmbSigla.List(cmbSigla.ListIndex) & "')"
    End If
    If cmbSiglaCh.ListIndex > 0 Then
       qryExe = qryExe & " AND ( SIGLACH ='" & cmbSiglaCh.List(cmbSiglaCh.ListIndex) & "')"
    End If
    AggiornaGriglia
End Sub

Private Sub cmdFiltraAna_Click()
Screen.MousePointer = vbHourglass
qryExe = Query
qryExe = qryExe & " WHERE 1=1" & Filtro

    If TxtRicCodAvvInt.Text <> "" Then
        qryExe = qryExe & " AND(AnagraficaAvvocati.CODAVV  LIKE '" & TxtRicCodAvvInt.Text & "%')"
    End If
    
    If TxtRicNome.Text <> "" Then
        qryExe = qryExe & " AND(AnagraficaAvvocati.NOME Like '" & Replace(TxtRicNome.Text, "'", "''") & "%')"
    End If
    
    If TxtUsufruente.Text <> "" Then
        qryExe = "SELECT AnagraficaAvvocati.CODAVV as Codice, AnagraficaAvvocati.NOME as Nome,  AnagraficaAvvocati.Telef as Telefono,AnagraficaAvvocati.TelefCell as Cellulare,NumOrdinamento "
        qryExe = qryExe & "FROM AnagraficaAvvocati INNER JOIN Usufruenti ON AnagraficaAvvocati.CODAVV = Usufruenti.CODAVV"
        qryExe = qryExe & " WHERE (((Usufruenti.DescrizioneUsufr) Like '" & Replace(TxtUsufruente.Text, "'", "''") & "%'))"
    End If
    AggiornaGriglia
Screen.MousePointer = vbDefault
End Sub

Private Sub flex_AfterSort(ByVal Col As Long, Order As Integer)
ColoraAnnullati
If Col = 3 Then
  AggiornaGriglia
 Else
  sortGrid flex, Col, Order, 1, -1
End If
End Sub

Private Sub flex_DblClick()

Dim r As Long
r = flex.row

If Not frmCaller Is Nothing Then
  frmCaller.IsLoading = True
End If
If r < 1 Then Exit Sub
If tipo = "Anagrafica" Then
   If Not frmCaller Is Nothing Then
      If TypeOf frmCaller Is IAnagraficForm Then
        Dim iAnaForm As IAnagraficForm
        Set iAnaForm = frmCaller
        iAnaForm.RisultatoRicerca flex.TextMatrix(r, 1), TipoAzione.Nuovo
      End If
   End If
   mFoundCode = flex.TextMatrix(r, 1)
   RaiseEvent AvvocatoSelezionato(mFoundCode)
 Else
   If Not frmCaller Is Nothing Then
     
     
     frmCaller.Where = "IDCod= " & flex.TextMatrix(r, PosizioneCodice)
     frmCaller.RisRicerca
    
    
   End If
   mFoundCode = flex.TextMatrix(r, 1)
   RaiseEvent AvvocatoSelezionato(mFoundCode)
End If
 If Not frmCaller Is Nothing Then
    frmCaller.SetFocus
    
    frmCaller.IsLoading = False
 End If
  
'Unload Me
End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If flex.ColIndex("Nome") <> -1 Then
    If tipo = "Anagrafica" And Button = 2 And flex.TextMatrix(flex.row, flex.ColIndex("Nome")) = "" Then
      PopupMenu mnuContext
    End If
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Caption = Titolo
If tipo = "Anagrafica" Then
  Query = "SELECT CODAVV as Codice, NOME as Nome,Telef as Telefono, TelefCell as Cellulare,NumOrdinamento FROM AnagraficaAvvocati "
  DefaultOrder = "order by AnagraficaAvvocati.NumOrdinamento"
  fraAna.Visible = True
  fraTop.Visible = False
  NCol = 4  'Numero di colonne da visualizzare
 Else
  fraAna.Visible = False
  fraTop.Visible = True
  PopolaCombo cmbSigla, "SELECT DISTINCT SIGLA as s FROM " & Tabella, "s", , , True
  PopolaCombo cmbSiglaCh, "SELECT DISTINCT SIGLACH as s FROM " & Tabella, "s", , , True
  cmbDate.AddItem "Mese"
  cmbDate.AddItem "Trimestre"
  cmbDate.AddItem "Semestre"
  cmbDate.AddItem "Anno"
  cmbDate.AddItem "Anno " & year(Date)
  cmbDate.AddItem "Tutto"
  cmbDate.ListIndex = 5
End If
 qryExe = Query & " WHERE 1=1" & Filtro
 AggiornaGriglia
 Ridimensiona
If tipo = "Anagrafica" Then
  
Else
  flex.ColDataType(3) = flexDTDate
  
End If
Screen.MousePointer = vbDefault
End Sub
Public Sub Ridimensiona()
If Atap.ScaleHeight - Atap.Toolbar1.Height > 0 Then Me.Move Location, 0, Me.Width, Atap.ScaleHeight
If Me.Height - flex.Top - 600 > 0 Then flex.Height = Me.ScaleHeight - flex.Top - 600
cmdEsci.Top = flex.Top + flex.Height + 80
End Sub
Public Sub AggiornaGriglia()
Dim i As Integer
Dim rs As ADODB.Recordset
Set rs = newAdoRs
rs.Open qryExe & " " & DefaultOrder, g_Settings.DBConnection
Set flex.DataSource = rs
For i = NCol + 1 To flex.Cols - 1
 flex.ColHidden(i) = True
Next i

For i = 1 To flex.Cols - 1
 flex.ColWidth(i) = 20
Next i
flex.ColWidth(flex.ColIndex("Codice")) = 900
If tipo = "Anagrafica" Then
    flex.ColWidth(flex.ColIndex("Nome")) = 3200
    flex.ColWidth(flex.ColIndex("Telefono")) = 1900
    flex.ColWidth(flex.ColIndex("Cellulare")) = 1700
 Else
 
     flex.ColWidth(flex.ColIndex("Ev")) = 200
     flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1600
     flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1400
     flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1300
     
     flex.ColAlignment(flex.ColIndex("Data Registrazione")) = flexAlignCenterCenter
     flex.ColAlignment(flex.ColIndex("Sigla Inserimento")) = flexAlignCenterCenter
     flex.ColAlignment(flex.ColIndex("Sigla chiusura")) = flexAlignCenterCenter

    Select Case Tabella
      Case "ADEMPI"
       flex.ColWidth(flex.ColIndex("Attività")) = 7000
      Case "SFRATTI"
        flex.ColWidth(flex.ColIndex("Parte1")) = 3400
        flex.ColWidth(flex.ColIndex("Parte2")) = 3400
      Case "SFRATTI_UNEP"
      flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1500
        flex.ColWidth(flex.ColIndex("Parte1")) = 2800
        flex.ColWidth(flex.ColIndex("Parte2")) = 2800
        flex.ColWidth(flex.ColIndex("Crono")) = 2200
        flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1200
     flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1200
      Case "NOTIFICHE"
        flex.ColWidth(flex.ColIndex("Parte1")) = 3400
        flex.ColWidth(flex.ColIndex("Parte2")) = 3400
      Case "DecretiIngiuntivi"
        flex.ColWidth(flex.ColIndex("Ricorrente")) = 3400
        flex.ColWidth(flex.ColIndex("Debitore")) = 3400
      Case "NOTIFICHE_UNEP"
      flex.ColWidth(flex.ColIndex("Data Registrazione")) = 1500
        flex.ColWidth(flex.ColIndex("Parte1")) = 2800
        flex.ColWidth(flex.ColIndex("Parte2")) = 2800
        flex.ColWidth(flex.ColIndex("Crono")) = 2200
        flex.ColWidth(flex.ColIndex("Sigla Inserimento")) = 1200
        flex.ColWidth(flex.ColIndex("Sigla chiusura")) = 1200
    End Select
   
    
    
    
    
    'flex.ColWidth(flex.ColIndex("Parte1")) = 2000
    flex.ColWidth(0) = 200
    
    
 
End If
  ColoraAnnullati
End Sub

Private Sub ColoraAnnullati()
Dim i As Long
ColoraLiberi
ColoraEvasi
If flex.ColIndex("Annullo") = -1 Then Exit Sub
  For i = 1 To flex.Rows - 1
    
    If flex.TextMatrix(i, flex.ColIndex("Annullo")) = "A" Then
      flex.Cell(flexcpForeColor, i, 1, i, flex.Cols - 1) = &HC0C0C0
      flex.Cell(flexcpFontStrikethru, i, 1, i, flex.Cols - 1) = True
    End If
  Next i
 
End Sub

Private Sub ColoraEvasi()
Dim i As Long
If flex.ColIndex("Ev") = -1 Then Exit Sub
  For i = 1 To flex.Rows - 1
    
    If flex.TextMatrix(i, flex.ColIndex("Ev")) = "X" Then
      flex.Cell(flexcpForeColor, i, 1, i, flex.Cols - 1) = &HFF0000
      
    End If
  Next i
 
End Sub
Private Sub ColoraLiberi()
Dim i As Long
If flex.ColIndex("Telefono") = -1 Then Exit Sub
  For i = 1 To flex.Rows - 1
    
    If flex.TextMatrix(i, flex.ColIndex("Nome")) = "" Then
      flex.Cell(flexcpBackColor, i, 1, i, flex.Cols - 1) = &H80FF80
      
    End If
  Next i
 
End Sub


Private Sub mnuElimina_Click()
 Dim r As Long
 Dim codAvv As String
 codAvv = flex.TextMatrix(flex.row, 1)
 
 r = MsgBox("Sei sicuro di voler eliminare la cassetta " & codAvv)
 If r = vbOK Then
    g_Settings.DBConnection.Execute "DELETE * FROM AnagraficaAvvocati Where CodAvv='" & codAvv & "'"
    AggiornaGriglia
 End If
 
End Sub

Private Sub opt_Click(Index As Integer)
 cmdFiltra_Click
End Sub


Private Sub TxtRicCodAvv_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdFiltra_Click
End Sub
