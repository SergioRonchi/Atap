VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Atap 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   Caption         =   "ATAP Service 2.0"
   ClientHeight    =   5415
   ClientLeft      =   1425
   ClientTop       =   2160
   ClientWidth     =   16755
   Icon            =   "Atap.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrStorico 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6240
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "23/06/2016"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18785
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "BLOC MAIUSC"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "BLOC NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "10.35"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListSmall 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":030A
            Key             =   "Adempimenti"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":0464
            Key             =   "Sfratti"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":05BE
            Key             =   "Stampa"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":0718
            Key             =   "Decreti"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":0872
            Key             =   "Anagrafica"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":09CC
            Key             =   "Uscita"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":0E1E
            Key             =   "Notifiche"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":0F78
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":3972
            Key             =   "Storico"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":636C
            Key             =   "Corrente"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":8D66
            Key             =   "NotificheUNEP"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":91B8
            Key             =   "SfrattiUNEP"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":960A
            Key             =   "Adempimenti"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":9A5C
            Key             =   "Anagrafica"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":A0A6
            Key             =   "Uscita"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":A4F8
            Key             =   "Notifiche"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":A94A
            Key             =   "Stampa"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":AC64
            Key             =   "Sfratti"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":AF96
            Key             =   "Decreti"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":B3E8
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":DDE2
            Key             =   "Storico"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":107DC
            Key             =   "Corrente"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":131D6
            Key             =   "NotificheUNEP"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atap.frx":13628
            Key             =   "SfrattiUNEP"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Anagrafica"
            Object.ToolTipText     =   "Anagrafica Avvocati"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Adempimenti"
            Object.ToolTipText     =   "Adempimenti di cancelleria"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sfratti"
            Object.ToolTipText     =   "Sfratti e Pignoramenti"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Notifiche"
            Object.ToolTipText     =   "Notifiche"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Decreti"
            Object.ToolTipText     =   "Decreti ingiuntivi"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Stampa"
            Object.ToolTipText     =   "Stampe"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Storico"
            Object.ToolTipText     =   "Apri DB Storico"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Corrente"
            Object.ToolTipText     =   "Ritorna al DB Corrente"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Lock"
            Object.ToolTipText     =   "Sblocca Tabelle"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Uscita"
            Object.ToolTipText     =   "Esci"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "NotificheUNEP"
            Object.ToolTipText     =   "Notifiche Unep"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "SfrattiUNEP"
            Object.ToolTipText     =   "Pignoramenti Unep"
         EndProperty
      EndProperty
      Begin VB.PictureBox pctStorico 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   8640
         ScaleHeight     =   375
         ScaleWidth      =   3855
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label lblStorico 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "Storico"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   3855
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuApriCorrente 
         Caption         =   "&Apri Archivio Corrente"
      End
      Begin VB.Menu mnuApriStorico 
         Caption         =   "Apri Archivio &Storico"
      End
      Begin VB.Menu mnuApriLiquidazione 
         Caption         =   "Apri Archivio &Liquidazione"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsporta 
         Caption         =   "Esporta Contabilità ..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Esegui backup"
      End
      Begin VB.Menu mnuss2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "E&sci"
      End
   End
   Begin VB.Menu mnuModuli 
      Caption         =   "&Moduli"
      Begin VB.Menu mnuAnagAvvoc 
         Caption         =   "&Anagrafica Avvocati"
      End
      Begin VB.Menu mnuAdempCancel 
         Caption         =   "A&dempimenti di cancelleria"
      End
      Begin VB.Menu mnuSfrattiPignoramenti 
         Caption         =   "&Sfratti e Pignoramenti"
      End
      Begin VB.Menu mnuNotifiche 
         Caption         =   "&Notifiche"
      End
      Begin VB.Menu mnu_DecretiIngiuntivi 
         Caption         =   "&Decreti Ingiuntivi"
      End
      Begin VB.Menu mnuSepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_NotificheUNEP 
         Caption         =   "Notifiche UNEP"
      End
      Begin VB.Menu mnu_PignoramentiUNEP 
         Caption         =   "Pignoramenti UNEP"
      End
   End
   Begin VB.Menu mnuGestioneTabelle 
      Caption         =   "&Gestione Tabelle"
      Begin VB.Menu mnuPignoramenti 
         Caption         =   "&Pignoramenti"
      End
      Begin VB.Menu mnuTribunali 
         Caption         =   "&Tribunali"
      End
      Begin VB.Menu mnuAnticipi 
         Caption         =   "&Anticipi"
      End
      Begin VB.Menu mnuSaldi 
         Caption         =   "&Saldi"
      End
      Begin VB.Menu mnuProvince 
         Caption         =   "P&rovince"
      End
      Begin VB.Menu mnuSGT 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnepSaldi 
         Caption         =   "UNEP Saldi"
      End
   End
   Begin VB.Menu mnuStampe 
      Caption         =   "&Stampe"
      Begin VB.Menu mnu_StampaGGAttivita 
         Caption         =   "&Giornaliera Attività"
      End
      Begin VB.Menu mnuAnagAvvocati 
         Caption         =   "Anagrafica &Avvocati"
      End
      Begin VB.Menu mnuSep34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstrattoConto 
         Caption         =   "&Estratto Conto"
      End
      Begin VB.Menu mnuEstrattoContoAdempimenti 
         Caption         =   "E/&C Adempimenti"
      End
      Begin VB.Menu mnuSospesi 
         Caption         =   "Ine&vasi"
      End
      Begin VB.Menu mnuStampaFatture 
         Caption         =   "&Fatture"
      End
      Begin VB.Menu mnuStampaSaldiProv 
         Caption         =   "Saldi &Provvisori"
      End
      Begin VB.Menu mnuStampaAssCircProv 
         Caption         =   "Asseg&ni Circolari Provvisoria"
      End
      Begin VB.Menu mnuTabelle 
         Caption         =   "&Tabelle"
         Begin VB.Menu mnuStampePignoramenti 
            Caption         =   "&Pignoramenti"
         End
         Begin VB.Menu mnuStampeTribunali 
            Caption         =   "&Tribunali"
         End
         Begin VB.Menu mnuStampeAnticipi 
            Caption         =   "&Anticipi"
         End
         Begin VB.Menu mnuStampeSaldi 
            Caption         =   "&Saldi"
         End
      End
      Begin VB.Menu mnuS56 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnepEC 
         Caption         =   "&UNEP Estratto conto"
      End
      Begin VB.Menu mnuUnepInevasi 
         Caption         =   "UNEP Inevasi"
      End
      Begin VB.Menu mnuUneoFat 
         Caption         =   "UNEP Fattu&re"
      End
      Begin VB.Menu mnuUnepSaldiProv 
         Caption         =   "UNEP Saldi provvisori"
      End
      Begin VB.Menu mnuUNEPAssegni 
         Caption         =   "UNEP Assegni Circolari provvisori"
      End
      Begin VB.Menu mnuUNEPTabSaldi 
         Caption         =   "UNEP Saldi"
      End
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "S&trumenti"
      Begin VB.Menu mnuFattura 
         Caption         =   "&Genera Fattura"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "&Calcolatrice"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Opzioni"
      Begin VB.Menu mnuOpt 
         Caption         =   "Configura"
         Index           =   0
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Barra degli strumenti"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Icone grandi"
         Checked         =   -1  'True
         Index           =   2
      End
   End
   Begin VB.Menu mnuUtilita 
      Caption         =   "&Utilità"
      Begin VB.Menu mnuOrdinamentoAvvocati 
         Caption         =   "&Ordinamento Avvocati"
      End
      Begin VB.Menu mnuGestioneCambioCassetta 
         Caption         =   "&Gestione Cambio Cassetta"
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSblocca 
         Caption         =   "&Sblocca Stampe"
      End
      Begin VB.Menu mnuSbloccaTab 
         Caption         =   "Sblocca &Tabelle"
      End
   End
End
Attribute VB_Name = "Atap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
On Error Resume Next
 frmback.Show
 mnuOpt(1).Checked = GetSetting("ATAP", "Configura", "Toolbar", 1)
 mnuOpt(2).Checked = GetSetting("ATAP", "Configura", "IconBig", 1)
 CaricaIcone mnuOpt(2).Checked
 If Me.ScaleWidth - pctStorico.Width > 0 Then pctStorico.Move (Me.ScaleWidth - pctStorico.Width) / 2
 Toolbar1.Visible = mnuOpt(1).Checked
 statusBar.Panels(2).Text = g_Settings.dbFile
 Caption = "Atap Service v." & app.Major & "." & app.Minor & "." & app.Revision
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 Uscita
 
End Sub

Private Sub mnu_DecretiIngiuntivi_Click()
    frmDecretiIngiuntivi.Show
End Sub

Private Sub mnu_NotificheUNEP_Click()
Notifiche.isUnep = True
Notifiche.Show
End Sub

Private Sub mnu_PignoramentiUNEP_Click()
SfrattiPignoramenti.isUnep = True
SfrattiPignoramenti.Show
End Sub

Private Sub mnu_StampaGGAttivita_Click()
    StampaGiornalieraAttivita.Show
End Sub

Private Sub mnuAdempCancel_Click()
    AdempCancel.Show
End Sub

Public Sub mnuAnagAvvoc_Click()
    Set FrmRicerca.frmCaller = AnagAvvocati
    FrmRicerca.tipo = "Anagrafica"
    
    FrmRicerca.Titolo = "Anagrafica Avvocati"
    FrmRicerca.Filtro = ""
    If FindForm("frmRicerca") Then
          Unload FrmRicerca
    End If

    Load FrmRicerca
    
    'AnagAvvocati.Show
End Sub

Private Sub mnuAnagAvvocati_Click()
    StampaAnagraficaAvvocati.Show
End Sub

Private Sub mnuAnticipi_Click()
    Anticipi.Show
End Sub

Private Sub mnuApriCorrente_Click()
On Error GoTo fine
pctStorico.Visible = False
tmrStorico.Enabled = False
g_Settings.ConnettiDB (g_Settings.dbFile)

MsgBox "Database Corrente Aperto correttamente", vbInformation
Exit Sub
fine:
 MsgBox err.Description, vbCritical
End Sub

Private Sub mnuApriLiquidazione_Click()

frmApriStorico.Path = "Liquidazioni"
frmApriStorico.Show vbModal
If frmApriStorico.aperto Then
    CloseAllForms
    lblStorico.Caption = "Liquidazione " & frmApriStorico.codice
    pctStorico.Visible = True
    tmrStorico.Enabled = True
End If

End Sub

Private Sub mnuApriStorico_Click()

frmApriStorico.Path = "EstrattiConto"
frmApriStorico.Show vbModal
If frmApriStorico.aperto Then
    CloseAllForms
    lblStorico.Caption = "E.C. Storico"
    pctStorico.Visible = True
    tmrStorico.Enabled = True
End If
End Sub

Private Sub mnuBackup_Click()
 Dim fb As New FileBackuoHelper
 pctStorico.Visible = False
 tmrStorico.Enabled = False
 fb.BackUp g_Settings.AtapUserBackupFolder
End Sub

Private Sub mnuCalc_Click()
 Shell ("Calc.exe")
End Sub



Private Sub mnuEsci_Click()
Uscita
End Sub

Private Sub mnuEsporta_Click()

 Dim frm As frmEsportaProfis
 Set frm = New frmEsportaProfis
 frm.Show vbModal
End Sub

Private Sub mnuEstrattoConto_Click()
    StampaEstrattoConto.Show
End Sub

Private Sub mnuEstrattoContoAdempimenti_Click()
    StampaEstrattoContoAdempimenti.Show
End Sub

Private Sub mnuFattura_Click()
    GeneraFattura.Show
End Sub

Private Sub mnuGestioneCambioCassetta_Click()
    GestioneCambioCassetta.Show
End Sub



Private Sub mnuNotifiche_Click()
    Notifiche.isUnep = False
    Notifiche.Show
End Sub

Private Sub mnuOpt_Click(Index As Integer)
Select Case Index
  Case 0
   frmConfigurazione.Show vbModal
  Case 1
    mnuOpt(1).Checked = Not mnuOpt(1).Checked
    Me.Toolbar1.Visible = mnuOpt(1).Checked
  Case 2 'Icone grandi
    mnuOpt(2).Checked = Not mnuOpt(2).Checked
    CaricaIcone (mnuOpt(2).Checked)
 End Select
End Sub
Private Sub CaricaIcone(IconeGrandi As Boolean)
Dim i As Integer
  If IconeGrandi Then
     Me.Toolbar1.ImageList = ImageList1
     Toolbar1.ButtonWidth = 464
     Toolbar1.ButtonHeight = 464
     For i = 1 To Toolbar1.Buttons.count
       If Toolbar1.Buttons(i).Description <> "" Then
           Toolbar1.Buttons(i).Image = ImageList1.ListImages(Toolbar1.Buttons(i).Description).Index
       End If
     Next i
     
    Else
     Me.Toolbar1.ImageList = ImageListSmall
     Toolbar1.ButtonWidth = 300
     Toolbar1.ButtonHeight = 300
    For i = 1 To Toolbar1.Buttons.count
     If Toolbar1.Buttons(i).Description <> "" Then
      Toolbar1.Buttons(i).Image = ImageListSmall.ListImages(Toolbar1.Buttons(i).Description).Index
     End If
    Next i

    End If
    

End Sub
Private Sub mnuOrdinamentoAvvocati_Click()
    FrmOrdinamentoAvvocati.Show
End Sub



Private Sub mnuPignoramenti_Click()
    Pignoramenti.Tabella = "Pignoramenti"
    Pignoramenti.Campo1 = "Codice"
    Pignoramenti.Campo2 = "Descrizione"
    Pignoramenti.Ordinamento = "Descrizione"
    Pignoramenti.Show
End Sub

Private Sub mnuProvince_Click()
frmProvince.Show
End Sub

Private Sub mnuSaldi_Click()
    Saldi.Tabella = "Pignoramenti"
    Saldi.Campo1 = "Codice"
    Saldi.Campo2 = "Descrizione"
    Saldi.isUnep = False

    Saldi.Show
End Sub

Private Sub mnuSblocca_Click()
Dim r As Integer
r = MsgBox("Attenzione:" & vbCrLf & "Questa utility sblocca le stampe." & vbCrLf & _
         "Deve essere eseguita solo in caso di blocchi non motivati, " & vbCrLf & _
         "quando si è certi che nessuno sta stampando." & vbCrLf & _
         "Vuoi sbloccare le stampe ora?", vbQuestion + vbYesNo)
         
If r = vbYes Then DeLockAllPrtTables
End Sub

Private Sub mnuSbloccaTab_Click()
DeLockAllTables
End Sub

Private Sub mnuSfrattiPignoramenti_Click()
    SfrattiPignoramenti.isUnep = False
    SfrattiPignoramenti.Show
End Sub

Private Sub mnuSospesi_Click()
    StampaSospesi.Show
End Sub

Private Sub mnuStampaAssCircProv_Click()
    StampaAssCircProv.Show
End Sub

Private Sub mnuStampaFatture_Click()
    StampaFatture.Show
End Sub

Private Sub mnuStampaSaldiProv_Click()
    StampaSaldiNegativi.Show
End Sub

Private Sub mnuStampeAnticipi_Click()
    StampaAnticipi.Show
End Sub

Private Sub mnuStampePignoramenti_Click()
     StampaPignoramenti.Show
End Sub

Private Sub mnuStampeSaldi_Click()
    StampaSaldi.Show
End Sub

Private Sub mnuStampeTribunali_Click()
    StampaTribunali.Show
End Sub

Private Sub mnuTribunali_Click()
    Pignoramenti.Tabella = "TribunaliAppartenenza"
    Pignoramenti.Campo1 = "CodiceTribunale"
    Pignoramenti.Campo2 = "DescrizioneTribunale"
    Pignoramenti.Ordinamento = "DescrizioneTribunale"
    Pignoramenti.Show
    
End Sub


Private Sub mnuUneoFat_Click()
  StampaFattureUNEP.Show
End Sub

Private Sub mnuUNEPAssegni_Click()
 StampaAssCircProvUNEP.Show
End Sub

Private Sub mnuUnepEC_Click()
StampaEstrattoContoUNEP.Show
End Sub

Private Sub mnuUnepInevasi_Click()
StampaSospesiUNEP.Show
End Sub

Private Sub mnuUNEPSaldi_Click()
    Saldi.Tabella = "Pignoramenti"
    Saldi.Campo1 = "Codice"
    Saldi.Campo2 = "Descrizione"
    Saldi.isUnep = True

    Saldi.Show
End Sub

Private Sub mnuUnepSaldiProv_Click()
  StampaSaldiNegativiUNEP.Show
End Sub

Private Sub mnuUNEPTabSaldi_Click()
 StampaSaldiUNEP.Show
End Sub

Private Sub tmrStorico_Timer()
   
 If lblStorico.ForeColor = 8388608 Then
     lblStorico.ForeColor = 65535
   Else
     lblStorico.ForeColor = 8388608
 End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Description
Case "NotificheUNEP"
  mnu_NotificheUNEP_Click
Case "SfrattiUNEP"
  mnu_PignoramentiUNEP_Click
 Case "Anagrafica"
  mnuAnagAvvoc_Click
 Case "Adempimenti"
  mnuAdempCancel_Click
 Case "Sfratti"
  mnuSfrattiPignoramenti_Click
 Case "Notifiche"
  mnuNotifiche_Click
 Case "Decreti"
  mnu_DecretiIngiuntivi_Click
 Case "Stampa"
  PopupMenu mnuStampe
 Case "Uscita"
  Uscita
 Case "Storico"
  mnuApriStorico_Click
 Case "Corrente"
  mnuApriCorrente_Click
 Case "Lock"
  mnuSbloccaTab_Click
End Select

End Sub
Public Sub Uscita()
 SaveSetting "ATAP", "Configura", "Toolbar", CInt(mnuOpt(1).Checked)
 SaveSetting "ATAP", "Configura", "IconBig", CInt(mnuOpt(2).Checked)
 DeLockAllTables True
 End
End Sub
