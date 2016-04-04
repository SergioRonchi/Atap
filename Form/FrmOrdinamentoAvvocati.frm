VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form FrmOrdinamentoAvvocati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ordinamento Avvocati"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame FrmDatiAvvocato 
      Height          =   840
      Left            =   120
      TabIndex        =   3
      Top             =   45
      Width           =   7395
      Begin TDBNumber6Ctl.TDBNumber Num 
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calculator      =   "FrmOrdinamentoAvvocati.frx":0000
         Caption         =   "FrmOrdinamentoAvvocati.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmOrdinamentoAvvocati.frx":008C
         Keys            =   "FrmOrdinamentoAvvocati.frx":00AA
         Spin            =   "FrmOrdinamentoAvvocati.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   2011496453
         Value           =   0
         MaxValueVT      =   1127088133
         MinValueVT      =   775749637
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Height          =   690
         Left            =   6120
         TabIndex        =   4
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label LblDescrizione 
         Height          =   300
         Left            =   2760
         TabIndex        =   8
         Top             =   120
         Width           =   3300
      End
      Begin VB.Label LblCodAvvocato 
         Caption         =   "Cod.Cassetta :"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label LblCodice 
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   120
         Width           =   660
      End
      Begin VB.Label LblNumOrdinamento 
         Caption         =   "Num. Ordinamento :"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1545
      End
   End
   Begin VB.Frame FrmRicerca 
      Height          =   5115
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7395
      Begin VB.TextBox TxtRicerca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   2
         Top             =   120
         Width           =   1470
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "&Ricerca"
         Default         =   -1  'True
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid flex 
         Height          =   4575
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   7215
         _cx             =   12726
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
         ExplorerBar     =   3
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
End
Attribute VB_Name = "FrmOrdinamentoAvvocati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qry As String


Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub CmdRicerca_Click()
    Dim qry1 As String
    qry = "SELECT NumOrdinamento as Ord,CODAVV as Codice,NOME FROM AnagraficaAvvocati"
    If TxtRicerca.Text <> "" Then
        qry1 = "(CODAVV Like '" & TxtRicerca.Text & "%') OR NOME Like '" & TxtRicerca.Text & "%'"
        qry = qry + " WHERE " + qry1
    End If
    qry = qry + " ORDER BY NumOrdinamento "
    Call AggiornaGriglia(flex, qry)
    flex_DblClick
End Sub

Private Sub CmdSalva_Click()
On Error GoTo FINE
    Dim Response As String
    
    Dim cod As String
    
    If Num.Text = "" Then
        MsgBox "Numero Ordinamento obbligatorio!", vbInformation, "Attenzione"
        Num.SetFocus
        Exit Sub
    End If
    
    Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
        
        cod = UCase(LblCodice.Caption)
        g_Settings.DBConnection.BeginTrans
        g_Settings.DBConnection.Execute "UPDATE AnagraficaAvvocati SET NumOrdinamento=" & Num & _
                         " WHERE CODAVV='" & cod & "';"
        
        ' Aggiorna tabelle associate
        g_Settings.DBConnection.Execute "UPDATE Adempi SET NumOrdinamento=" & Num & _
                         " WHERE CODAVV='" & cod & "';"
        
        g_Settings.DBConnection.Execute "UPDATE Sfratti SET NumOrdinamento=" & Num & _
                         " WHERE CODAVV='" & cod & "';"
        
        g_Settings.DBConnection.Execute "UPDATE DecretiIngiuntivi SET NumOrdinamento=" & Num & _
                         " WHERE CODAVV='" & cod & "';"
                         
                         
        g_Settings.DBConnection.Execute "UPDATE Notifiche SET NumOrdinamento=" & Num & _
                         " WHERE CODAVV='" & cod & "';"
                         
        g_Settings.DBConnection.Execute "UPDATE Saldi SET NumOrdinamento=" & Num & _
                         " WHERE Codice='" & cod & "';"
                         
        g_Settings.DBConnection.CommitTrans
        MsgBox "Record Modificato!", vbInformation, "Informazione"
        AggiornaGriglia flex, qry
    End If
    
Exit Sub
FINE:
MsgBox err.Description
g_Settings.DBConnection.RollbackTrans

End Sub


Private Sub flex_DblClick()
Dim r As Integer
r = flex.row
LblCodice = flex.TextMatrix(r, 2)
LblDescrizione = flex.TextMatrix(r, 3)
Num = flex.TextMatrix(r, 1)
CmdSalva.Enabled = True
End Sub

Private Sub Form_Load()

qry = "SELECT NumOrdinamento as Ord,CODAVV as Codice,NOME FROM AnagraficaAvvocati Order By NumOrdinamento"
Call AggiornaGriglia(flex, qry)
flex_DblClick

End Sub

