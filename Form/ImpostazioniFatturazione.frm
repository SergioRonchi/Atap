VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form ImpostazioniFatturazione 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impostazioni per fatturazione"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalva 
      Caption         =   "&Salva"
      Height          =   555
      Left            =   2640
      TabIndex        =   2
      Top             =   1020
      Width           =   915
   End
   Begin TDBDate6Ctl.TDBDate txtData 
      DataField       =   "DataRegistrazione"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Tag             =   "necessario Data Registrazione"
      Top             =   120
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calendar        =   "ImpostazioniFatturazione.frx":0000
      Caption         =   "ImpostazioniFatturazione.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "ImpostazioniFatturazione.frx":0184
      Keys            =   "ImpostazioniFatturazione.frx":01A2
      Spin            =   "ImpostazioniFatturazione.frx":0200
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
   Begin TDBNumber6Ctl.TDBNumber txtNumero 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   503
      Calculator      =   "ImpostazioniFatturazione.frx":0228
      Caption         =   "ImpostazioniFatturazione.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "ImpostazioniFatturazione.frx":02B4
      Keys            =   "ImpostazioniFatturazione.frx":02D2
      Spin            =   "ImpostazioniFatturazione.frx":031C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "######0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "######0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999
      MinValue        =   0
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
   Begin VB.Label LblNumero 
      Caption         =   "Numero Fattura Iniziale :"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   585
      Width           =   1770
   End
   Begin VB.Label LblData 
      Caption         =   "Data :"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   1680
   End
End
Attribute VB_Name = "ImpostazioniFatturazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isUnep As Boolean
Private Sub CmdSalva_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    txtData = Date
    txtNumero = getNewNumFattura
    If isUnep Then
      StampaEstrattoContoUNEP.Hide
    Else
      StampaEstrattoConto.Hide
    End If
    
    Atap.mnuModuli.Enabled = False
    Atap.mnuGestioneTabelle.Enabled = False
    Atap.mnuStampe.Enabled = False
    Atap.mnuStrumenti.Enabled = False
    Atap.mnuUtilita.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error GoTo fine
Dim dataFattura As ADODB.Recordset
  Screen.MousePointer = vbHourglass
    If isUnep Then
        Set dataFattura = GetADORecordset("Date_EstrattiConto", "DATA_FATTURA_UNEP", "1=1", g_Settings.DBConnection)
        If dataFattura Is Nothing Then
             g_Settings.DBConnection.Execute ("INSERT INTO Date_EstrattiConto (DATA_FATTURA_UNEP) VALUES ('" & txtData & "')")
            Else
             g_Settings.DBConnection.Execute ("UPDATE Date_EstrattiConto SET DATA_FATTURA_UNEP= '" & txtData & "'")
        End If
        StampaEstrattoContoUNEP.GeneraFattura txtNumero, txtData.Text, False
        StampaEstrattoContoUNEP.Show
    Else
        Set dataFattura = GetADORecordset("Date_EstrattiConto", "DATA_FATTURA", "1=1", g_Settings.DBConnection)
        If dataFattura Is Nothing Then
             g_Settings.DBConnection.Execute ("INSERT INTO Date_EstrattiConto (DATA_FATTURA) VALUES ('" & txtData & "')")
            Else
             g_Settings.DBConnection.Execute ("UPDATE Date_EstrattiConto SET DATA_FATTURA= '" & txtData & "'")
        End If
        StampaEstrattoConto.GeneraFattura txtNumero, txtData.Text, False
        StampaEstrattoConto.Show
    End If

    
    
    Screen.MousePointer = vbDefault
    

    Atap.mnuModuli.Enabled = True
    Atap.mnuGestioneTabelle.Enabled = True
    Atap.mnuStampe.Enabled = True
    Atap.mnuStrumenti.Enabled = True
    Atap.mnuUtilita.Enabled = True
    MsgBox "Procedura di fatturazione eseguita con successo", vbInformation, "Informazione"
 Exit Sub
fine:
 MsgBox err.Description
 
End Sub




