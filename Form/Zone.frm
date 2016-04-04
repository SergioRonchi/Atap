VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Pignoramenti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pignoramenti"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   15
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   450
      Left            =   5280
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton CmdElimina 
      Caption         =   "&Elimina"
      Enabled         =   0   'False
      Height          =   450
      Left            =   4080
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame FrmRicerca 
      Height          =   4140
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6360
      Begin VB.TextBox TxtRicerca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   3135
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "&Ricerca"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid flex 
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   6015
         _cx             =   10610
         _cy             =   6165
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
   Begin VB.Frame FrmPignoramenti 
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Nuovo"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Salva"
         Height          =   690
         Left            =   5040
         TabIndex        =   9
         Top             =   120
         Width           =   1000
      End
      Begin VB.TextBox TxtCodice 
         DataField       =   "Codice"
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox TxtDescrizione 
         DataField       =   "Descrizione"
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   765
         Width           =   3165
      End
      Begin VB.Label LblCodice 
         Caption         =   "Codice"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   960
      End
      Begin VB.Label LblDescrizione 
         Caption         =   "Descrizione"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   765
         Width           =   960
      End
   End
End
Attribute VB_Name = "Pignoramenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qry As String
Public Tabella As String
Public Campo1 As String
Public Campo2 As String
Public Ordinamento As String



Private Sub CmdElimina_Click()
Dim app As String
Dim err As String
Dim Response As Long
app = flex.TextMatrix(flex.row, 1)
If Tabella = "TribunaliAppartenenza" Then
    err = ControllaTribunale(app)
    If err <> "" Then
      MsgBox "Non può essere eliminato il tribunale " & TxtDescrizione & " perché esistono voci in relazione." & vbCrLf & _
             "In particolare esistono:" & vbCrLf & err, vbOKOnly + vbInformation, "ATAP"
             
      Exit Sub
    End If
End If
Response = MsgBox("Vuoi eliminare il record  " & app & " ?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
If Response = vbYes Then    ' User chose Yes.
 g_Settings.DBConnection.Execute "DELETE FROM " & Tabella & " WHERE " & Campo1 & "='" & app & "'"
 If Tabella = "Pignoramenti" Then
   If FindForm("SfrattiPignoramenti") Then PopolaCombo SfrattiPignoramenti.CmbPignoramenti, "Pignoramenti", "Descrizione"
  Else
     If FindForm("AdempCancel") Then PopolaCombo AdempCancel.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

     If FindForm("SfrattiPignoramenti") Then PopolaCombo SfrattiPignoramenti.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

     If FindForm("Notifiche") Then PopolaCombo Notifiche.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

 End If
 Aggiorna
End If
End Sub

Private Sub cmdEsci_Click()
Unload Me
If FindForm("frmRicerca") Then
    Unload FrmRicerca
End If

End Sub

Private Sub cmdNew_Click()
On Error GoTo FINE
 
 If TxtCodice <> "" And TxtDescrizione <> "" Then
   If ExistADORecord("SELECT * From " & Tabella & " Where " & Campo1 & "='" & Replace(TxtCodice, "'", "''") & "'", g_Settings.DBConnection) Then
       g_Settings.DBConnection.Execute "UPDATE " & Tabella & " Set " & Campo1 & "='" & Replace(TxtCodice, "'", "''") & "', " & Campo2 & "='" & Replace(TxtDescrizione, "'", "''") & "' Where " & Campo1 & "='" & Replace(TxtCodice, "'", "''") & "'"
    Else
       g_Settings.DBConnection.Execute "INSERT INTO " & Tabella & " (" & Campo1 & "," & Campo2 & ") VALUES ('" & Replace(TxtCodice, "'", "''") & "', '" & Replace(TxtDescrizione, "'", "''") & "')"
   End If
     Aggiorna
     If Tabella = "Pignoramenti" Then
      If FindForm("SfrattiPignoramenti") Then PopolaCombo SfrattiPignoramenti.CmbPignoramenti, "Pignoramenti", "Descrizione"
     Else
     If FindForm("AdempCancel") Then PopolaCombo AdempCancel.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

     If FindForm("SfrattiPignoramenti") Then PopolaCombo SfrattiPignoramenti.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

     If FindForm("Notifiche") Then PopolaCombo Notifiche.cmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

 End If

   Else
   MsgBox "E' indispensabile inserire sia il codice che la descrizione.", vbOKOnly + vbCritical
   
 End If
 Exit Sub
FINE:
 MsgBox err.Description

End Sub

Private Sub CmdRicerca_Click()
  Dim qry1 As String
    qry = "SELECT " & Campo1 & "," & Campo2 & " FROM " & Tabella & ""
    If TxtRicerca.Text <> "" Then
        qry1 = "(" & Campo2 & " Like '" & TxtRicerca.Text & "%')"
        qry = qry + " WHERE " + qry1
    End If
   Aggiorna
         
End Sub

Private Sub Command1_Click()
PulisciCampi Me
End Sub

Private Sub flex_AfterSort(ByVal Col As Long, Order As Integer)
flex_DblClick
End Sub

Private Sub flex_DblClick()
Dim r As Long
r = flex.row
If r = 0 Then Exit Sub
TxtCodice = flex.TextMatrix(r, 1)
TxtDescrizione = flex.TextMatrix(r, 2)

CmdElimina.Enabled = TxtCodice <> "UNEP"
cmdNew.Enabled = TxtCodice <> "UNEP"
TxtCodice.Enabled = TxtCodice <> "UNEP"
TxtDescrizione.Enabled = TxtCodice <> "UNEP"
Command1.Enabled = TxtCodice <> "UNEP"
End Sub

Private Sub Form_Load()
 qry = "SELECT * FROM " & Tabella & ""
 If Ordinamento <> "" Then qry = qry & " ORDER BY " & Ordinamento
 Caption = Tabella
  Aggiorna
End Sub
Public Sub Aggiorna()
 AggiornaGriglia flex, qry, CmdElimina
 flex.ColWidth(2) = 3000
 flex_DblClick
End Sub

