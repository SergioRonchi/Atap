VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Tribunali 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tribunali"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmTribunali 
      Enabled         =   0   'False
      Height          =   1065
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   5775
      Begin VB.TextBox TxtDescrizione 
         Height          =   285
         Left            =   1725
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   3165
      End
      Begin VB.TextBox TxtCodice 
         Height          =   285
         Left            =   1725
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblDescrizione 
         Caption         =   "Descrizione"
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   960
      End
      Begin VB.Label LblCodice 
         Caption         =   "Codice"
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame FrmButtonZone 
      Height          =   1005
      Left            =   480
      TabIndex        =   10
      Top             =   4320
      Width           =   5280
      Begin VB.CommandButton CmdModifica 
         Caption         =   "&Modifica"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Width           =   1095
      End
      Begin VB.CommandButton CmdAggiungi 
         Caption         =   "&Aggiungi"
         Height          =   330
         Left            =   1305
         TabIndex        =   7
         Top             =   330
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalva 
         Caption         =   "&Salva"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2490
         TabIndex        =   8
         Top             =   345
         Width           =   1095
      End
      Begin VB.CommandButton CmdElimina 
         BackColor       =   &H000000FF&
         Caption         =   "&Elimina"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   345
         Width           =   1095
      End
   End
   Begin VB.Frame FrmRicercaTribunali 
      Height          =   3180
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5760
      Begin VB.Data DataRicerca 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   1395
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2070
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "&Ricerca"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtRicerca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   3135
      End
      Begin MSDBGrid.DBGrid DBGridRicercaTribunali 
         Bindings        =   "Tribunali.frx":0000
         Height          =   2595
         Left            =   120
         OleObjectBlob   =   "Tribunali.frx":001A
         TabIndex        =   9
         Top             =   480
         Width           =   5490
      End
   End
End
Attribute VB_Name = "Tribunali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GestioneTribunali As Recordset

Private Sub CmdElimina_Click()
Dim app As Variant
Dim err As String
app = TxtCodice.Text
err = ControllaTribunale(app)
If err <> "" Then
  MsgBox "Non può essere eliminato il tribunale " & TxtDescrizione & " perché esistono voci in relazione." & vbCrLf & _
         "In particolare esistono:" & vbCrLf & err, vbOKOnly + vbInformation, "ATAP"
         
  Exit Sub
End If
Response = MsgBox("Vuoi eliminare il tribunale " & TxtDescrizione & " ?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
If Response = vbYes Then    ' User chose Yes.
With GestioneTribunali
        .Index = "PrimaryKey"
        .Seek "=", app
        If .NoMatch Then
            MsgBox "Nessun tribunale " & app & "!"
        Else
            .Delete
            MsgBox "Tribunale " & app & " eliminato!"
            If IsActiveForm("AdempCancel") Then
                PopolaCombo AdempCancel.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

                
            End If
            If IsActiveForm("SfrattiPignoramenti") Then
               PopolaCombo SfrattiPignoramenti.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

                
            End If
            If IsActiveForm("Notifiche") Then
                PopolaCombo Notifiche.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

                
            End If
            'DataRicerca.Refresh
             refreshDbGridTribunali
            
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'Errore se la finestra AnagAvvocati non è caricata
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                'AnagAvvocati.RiempiCmbTribunali
                
        End If
         pulisci
 End With
 End If
End Sub
Private Function ControllaTribunale(cod) As String
  ControllaTribunale = ""
  If ExistRecord("SELECT CodTribunaleApp FROM ADEMPI Where CodTribunaleApp='" & cod & "';", gDB) Then
      ControllaTribunale = ControllaTribunale & " - Adempimenti;" & vbCrLf
      
  End If
  
  If ExistRecord("SELECT CodTribunaleApp FROM DecretiIngiuntivi Where CodTribunaleApp='" & cod & "';", gDB) Then
     ControllaTribunale = ControllaTribunale & " - Decreti Ingiuntivi;" & vbCrLf
      Exit Function
  End If
  
  If ExistRecord("SELECT CodTribunaleApp FROM Notifiche Where CodTribunaleApp='" & cod & "';", gDB) Then
            ControllaTribunale = ControllaTribunale & " - Notifiche;" & vbCrLf
      Exit Function
  End If
  
  If ExistRecord("SELECT CodTribunaleApp FROM Sfratti Where CodTribunaleApp='" & cod & "';", gDB) Then
            ControllaTribunale = ControllaTribunale & " - Sfratti;" & vbCrLf
      Exit Function
  End If
  
  If ExistRecord("SELECT CodiceTribunale FROM Anticipi Where CodiceTribunale='" & cod & "';", gDB) Then
            ControllaTribunale = ControllaTribunale & " - Anticipi;" & vbCrLf
      Exit Function
  End If
  
  If ExistRecord("SELECT CodTribunaleApp FROM ADEMPI Where CodTribunaleApp='" & cod & "';", gDB) Then
      ControllaTribunale = True
      Exit Function
  End If
End Function
Private Sub CmdRicerca_Click()
    Dim qry As String
     
    qry = "SELECT CodiceTribunale,DescrizioneTribunale FROM TribunaliAppartenenza"
    If TxtRicerca.Text <> "" Then
        qry1 = "(DescrizioneTribunale Like '" & TxtRicerca.Text & "%')"
        qry = qry + " WHERE " + qry1
    End If
    
    DataRicerca.DatabaseName = gDbName
    DataRicerca.RecordSource = qry
    refreshDbGridTribunali
         
End Sub

Private Sub CmdSalva_Click()
Dim Response As Variant

If TxtCodice.Text = "" Then
    MsgBox "Codice tribunale obbligatorio!", vbInformation, "Attenzione"
    TxtCodice.SetFocus
    Exit Sub
End If

If TxtDescrizione.Text = "" Then
    MsgBox "Descrizione tribunale obbligatoria!", vbInformation, "Attenzione"
    TxtDescrizione.SetFocus
    Exit Sub
End If

    GestioneTribunali.Index = "SecondaryKey"
    GestioneTribunali.Seek "=", UCase(TxtDescrizione.Text)
    If Not GestioneTribunali.NoMatch Then
        MsgBox "Tribunale già esistente!", vbInformation, "Attenzione"
        Exit Sub
    End If
 
If CmdAggiungi.Enabled = False Then
    'Sto Modificando la mia anagrafica
    Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
        GestioneTribunali.Index = "PrimaryKey"
        GestioneTribunali.Seek "=", UCase(TxtCodice.Text)
        GestioneTribunali.Edit
        RiempiRecordset
        GestioneTribunali.Update
        'MsgBox "Record Modificato!", vbInformation, "Informazione"
        If IsActiveForm("AdempCancel") Then
            PopolaCombo AdempCancel.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

            
        End If
        If IsActiveForm("SfrattiPignoramenti") Then
            PopolaCombo SfrattiPignoramenti.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

        End If
        If IsActiveForm("Notifiche") Then
            PopolaCombo Notifiche.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

        End If
        refreshDbGridTribunali
    End If
    CmdModifica.Caption = "&Modifica"
    CmdRicerca.Enabled = True
    CmdAggiungi.Enabled = True
    CmdSalva.Enabled = False
Else
    'Sto Aggiungendo un record alla mia anagrafica
    Response = MsgBox("Vuoi salvare i dati inseriti?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
    If Response = vbYes Then    ' User chose Yes.
        GestioneTribunali.Index = "SecondaryKey"
        GestioneTribunali.Seek "=", UCase(TxtDescrizione.Text)
        If Not GestioneTribunali.NoMatch Then
            MsgBox "Tribunale già esistente!", vbInformation, "Attenzione"
            Exit Sub
        End If
        GestioneTribunali.Index = "PrimaryKey"
        GestioneTribunali.Seek "=", UCase(TxtCodice.Text)
        If Not GestioneTribunali.NoMatch Then
            MsgBox "Codice Tribunale già esistente!", vbInformation, "Attenzione"
            Exit Sub
        End If
        GestioneTribunali.AddNew
        RiempiRecordset
        GestioneTribunali.Update
        'MsgBox "Record Salvato!", vbInformation, "Informazione"
        If IsActiveForm("AdempCancel") Then
            PopolaCombo AdempCancel.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"
            
        End If
        If IsActiveForm("SfrattiPignoramenti") Then
             PopolaCombo SfrattiPignoramenti.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"
        End If
        If IsActiveForm("Notifiche") Then
           PopolaCombo Notifiche.CmbTribunale, "TribunaliAppartenenza", "DescrizioneTribunale"

            
        End If
        GestioneTribunali.Index = "SecondaryKey"
        GestioneTribunali.Seek "=", UCase(TxtDescrizione.Text)
        refreshDbGridTribunali
    Else
        RiempiCampi
    End If
    CmdAggiungi.Caption = "&Aggiungi"
    CmdRicerca.Enabled = True
    CmdModifica.Enabled = True
    CmdSalva.Enabled = False
  End If
   pulisci
  
  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  'Errore se la finestra AnagAvvocati non è caricata
  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'AnagAvvocati.RiempiCmbTribunali
  
End Sub

Private Sub DBGridRicercaTribunali_DblClick()
    Dim Col0 As Column
    GestioneTribunali.Index = "PrimaryKey"
    Set Col0 = DBGridRicercaTribunali.Columns(0)
    GestioneTribunali.Seek "=", Col0.Text
    RiempiCampi
    CmdModifica.Enabled = True
    CmdElimina.Enabled = True
    CmdAggiungi.Enabled = True
End Sub

Private Sub Form_Load()
    Set GestioneTribunali = gDB.OpenRecordset("TribunaliAppartenenza", dbOpenTable)
    Call SetActiveForm("Tribunali")
 End Sub

Private Sub RiempiCampi()
    TxtCodice.Text = GestioneTribunali!CodiceTribunale
    TxtDescrizione.Text = GestioneTribunali!DescrizioneTribunale
End Sub
Private Sub CmdNext_Click()
    
    GestioneTribunali.MoveNext
    If Not GestioneTribunali.EOF Then
         RiempiCampi
    Else
        MsgBox "Sei all'ultimo record", vbInformation, "Attenzione"
        GestioneTribunali.MovePrevious
         RiempiCampi
    End If

End Sub

Private Sub CmdPrevious_Click()
   
    GestioneTribunali.MovePrevious
    If Not GestioneTribunali.BOF Then
         RiempiCampi
    Else
        MsgBox "Sei al primo record", vbInformation, "Attenzione"
        GestioneTribunali.MoveNext
         RiempiCampi
    End If

End Sub

Private Sub CmdFirst_Click()
    
    If GestioneTribunali.RecordCount > 0 Then
        GestioneTribunali.MoveFirst
             RiempiCampi
    End If

End Sub

Private Sub CmdLast_Click()
      
    If GestioneTribunali.RecordCount > 0 Then
        GestioneTribunali.MoveLast
             RiempiCampi
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    GestioneTribunali.Close
     Call UnloadActiveForm("Tribunali")
End Sub
Private Sub CmdModifica_Click()
    
    If CmdModifica.Caption = "&Modifica" Then
        FrmTribunali.Enabled = True
        TxtCodice.Enabled = False
        CmdElimina.Enabled = False
        CmdModifica.Caption = "&Annulla"
        FrmRicercaTribunali.Enabled = False
        CmdAggiungi.Enabled = False
        CmdSalva.Enabled = True
        TxtDescrizione.Enabled = True
        TxtDescrizione.SetFocus
    Else
        FrmTribunali.Enabled = False
        CmdModifica.Caption = "&Modifica"
        FrmRicercaTribunali.Enabled = True
        CmdAggiungi.Enabled = True
        CmdElimina.Enabled = True
        CmdSalva.Enabled = False
         pulisci
    End If

End Sub
Private Sub CmdAggiungi_Click()
    pulisci
    If CmdAggiungi.Caption = "&Aggiungi" Then
        FrmTribunali.Enabled = True
        TxtCodice.Enabled = True
        TxtDescrizione.Enabled = True
        TxtCodice.SetFocus
        CmdAggiungi.Caption = "&Annulla"
        FrmRicercaTribunali.Enabled = False
        CmdModifica.Enabled = False
        CmdSalva.Enabled = True
        CmdElimina.Enabled = False
       
     Else
        FrmTribunali.Enabled = False
        GestioneTribunali.Index = "PrimaryKey"
        CmdAggiungi.Caption = "&Aggiungi"
        FrmRicercaTribunali.Enabled = True
        CmdModifica.Enabled = True
        CmdSalva.Enabled = False
        pulisci
    End If

End Sub
Public Sub RiempiRecordset()
   GestioneTribunali!CodiceTribunale = TxtCodice.Text
   GestioneTribunali!DescrizioneTribunale = UCase(TxtDescrizione.Text)
End Sub

Private Sub pulisci()
   ' Ripristino situazione TxtField
    TxtDescrizione.Text = ""
    TxtCodice.Text = ""
    ' Ripristino situazione Frame
    FrmTribunali.Enabled = False
    FrmRicercaTribunali.Enabled = True
    ' Ripristino situazione Btn
    CmdElimina.Enabled = False
    CmdModifica.Enabled = False
    CmdAggiungi.Enabled = True
End Sub

Private Sub setWidthColDbGridTribunale()
    Dim col1, col2 As Column
    Set col1 = DBGridRicercaTribunali.Columns(0)
    Set col2 = DBGridRicercaTribunali.Columns(1)
    col1.Caption = "Codice"
    col2.Caption = "Descrizione"
    col1.Width = 600
    col2.Width = 3250
    DBGridRicercaTribunali.AllowRowSizing = False
End Sub

Public Sub refreshDbGridTribunali()
    DataRicerca.Refresh
    setWidthColDbGridTribunale
End Sub
