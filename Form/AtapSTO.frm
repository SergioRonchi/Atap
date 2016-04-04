VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm Atap 
   BackColor       =   &H8000000C&
   Caption         =   "Storico ATAP "
   ClientHeight    =   5475
   ClientLeft      =   1425
   ClientTop       =   2160
   ClientWidth     =   10665
   Icon            =   "AtapSTO.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Anagrafica Avvocati"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Adempimenti di Cancelleria"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Sfratti e Pignoramenti"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Notifiche"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Decreti Ingiuntivi"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Stampe"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Esci da Atap"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Height          =   570
         Left            =   5355
         ScaleHeight     =   510
         ScaleWidth      =   4245
         TabIndex        =   1
         Top             =   60
         Width           =   4305
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "STORICO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   15
            TabIndex        =   2
            Top             =   30
            Width           =   4245
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":030A
            Key             =   "Notifiche"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":0624
            Key             =   "Anagrafica"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":093E
            Key             =   "Stampe"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":0C58
            Key             =   "Sfratti"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":0F72
            Key             =   "Adempimenti"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":128C
            Key             =   "Decreti"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AtapSTO.frx":15A6
            Key             =   ""
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuParametri 
         Caption         =   "Pa&rametri"
      End
      Begin VB.Menu mnuSaldi 
         Caption         =   "&Saldi"
      End
   End
   Begin VB.Menu mnuStampe 
      Caption         =   "&Stampe"
      Begin VB.Menu mnu_StampaGGAttivita 
         Caption         =   "&Stampa Giornaliera Attività"
      End
      Begin VB.Menu mnuAnagAvvocati 
         Caption         =   "Stampa Anagrafica &Avvocati"
      End
      Begin VB.Menu mnuEstrattoConto 
         Caption         =   "Stampa &Estratto Conto"
      End
      Begin VB.Menu mnuEstrattoContoAdempimenti 
         Caption         =   "Stampa E/&C Adempimenti"
      End
      Begin VB.Menu mnuSospesi 
         Caption         =   "Stampa Ine&vasi"
      End
      Begin VB.Menu mnuStampaFatture 
         Caption         =   "Stampa &Fatture"
      End
      Begin VB.Menu mnuStampaSaldiProv 
         Caption         =   "S&tampa Saldi Provvisori"
      End
      Begin VB.Menu mnuStampaAssCircProv 
         Caption         =   "Sta&mpa Assegni Circolari Provvisoria"
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
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "S&trumenti"
      Begin VB.Menu mnuFattura 
         Caption         =   "&Genera Fattura"
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
   End
   Begin VB.Menu mnuUtilita 
      Caption         =   "&Utilità"
      Begin VB.Menu mnuOrdinamentoAvvocati 
         Caption         =   "&Ordinamento Avvocati"
      End
      Begin VB.Menu mnuGestioneCambioCassetta 
         Caption         =   "&Gestione Cambio Cassetta"
      End
      Begin VB.Menu mnuCompattaDB 
         Caption         =   "&Compatta Database"
      End
   End
End
Attribute VB_Name = "Atap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
 mnuOpt(1).Checked = GetSetting("ATAP", "Configura", "Toolbar", 1)
 Toolbar1.Visible = mnuOpt(1).Checked
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 Uscita
End Sub

Private Sub mnu_DecretiIngiuntivi_Click()
    DecretiIngiuntivi.Show
End Sub

Private Sub mnu_StampaGGAttivita_Click()
    StampaGiornalieraAttivita.Show
End Sub

Private Sub mnuAdempCancel_Click()
    AdempCancel.Show
End Sub

Private Sub mnuAnagAvvoc_Click()
    AnagAvvocati.Show
End Sub

Private Sub mnuAnagAvvocati_Click()
    StampaAnagraficaAvvocati.Show
End Sub

Private Sub mnuAnticipi_Click()
    Anticipi.Show
End Sub

Private Sub mnuCompattaDB_Click()
    Call CompattaDB
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
    Notifiche.Show
End Sub

Private Sub mnuOpt_Click(Index As Integer)
Select Case Index
  Case 0
  Case 1
    mnuOpt(1).Checked = Not mnuOpt(1).Checked
    Me.Toolbar1.Visible = mnuOpt(1).Checked
 End Select
End Sub

Private Sub mnuOrdinamentoAvvocati_Click()
    FrmOrdinamentoAvvocati.Show
End Sub

Private Sub mnuParametri_Click()
    Parametri.Show
End Sub

Private Sub mnuPignoramenti_Click()
    Pignoramenti.Show
End Sub

Private Sub mnuSaldi_Click()
    Saldi.Show
End Sub

Private Sub mnuSfrattiPignoramenti_Click()
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
    Tribunali.Show
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
 Case 1
  mnuAnagAvvoc_Click
 Case 2
  mnuAdempCancel_Click
 Case 3
  mnuSfrattiPignoramenti_Click
 Case 4
  mnuNotifiche_Click
 Case 5
  mnu_DecretiIngiuntivi_Click
 Case 7
  PopupMenu mnuStampe
 Case 9
  Uscita
End Select
End Sub
Public Sub Uscita()
 SaveSetting "ATAP", "Configura", "Toolbar", mnuOpt(1).Checked
 End
End Sub
