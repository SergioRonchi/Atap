VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Stampa 
   Caption         =   "Stampa ..."
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "StampaGiornaliera.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   Begin Crystal.CrystalReport crpt 
      Left            =   1980
      Top             =   1485
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowHeight    =   500
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "Stampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Event StampaEseguita(table As String)
Private m_Closed As Boolean
Private m_destination As DestinationConstants



Public Sub gestioneReport(TabellaStampa As String, qryFormula As String, tappo As Long, cDestination As DestinationConstants, sRptFile As String, Copie As Integer, Optional Formula0 As String)
Dim subRpt
On Error GoTo fine
  Form_Resize
   Screen.MousePointer = vbHourglass
   crpt.WindowParentHandle = Me.hwnd
   m_destination = cDestination
   
    With crpt
        
        .DataFiles(0) = g_Settings.CurrentDbFile
        '.DataFiles(1) = g_Settings.CurrentDbFile
        If sRptFile = "Fattura.rpt" Then
          .DataFiles(1) = g_Settings.CurrentDbFile
          .DataFiles(2) = g_Settings.CurrentDbFile
        End If

        .ReportFileName = g_Settings.ReportPath & "\" & sRptFile
        .SelectionFormula = Trim(qryFormula)
        .FetchSelectionFormula
        .WindowParentHandle = Me.hwnd
        If m_destination = crptToPrinter Then
           .CopiesToPrinter = Copie
        End If
        .WindowState = crptMaximized
        .Destination = m_destination
        If Formula0 <> "" Then
           .Formulas(0) = Formula0
          Else
           .Formulas(0) = ""
        End If
        If sRptFile = "EstrattoContoUNEP.rpt" Or sRptFile = "AnagraficaDettagliata.rpt" Then
          Dim n As Integer
          Dim sSubreportName As String
          n = .GetNSubreports
          If n = 1 Then
            sSubreportName = .GetNthSubreportName(0)
            .SubreportToChange = sSubreportName
            .DataFiles(0) = g_Settings.CurrentDbFile
            .SubreportToChange = ""
          End If
        End If
        .Action = 1
        RaiseEvent StampaEseguita(TabellaStampa)
       '.PrintReport
    End With
    Screen.MousePointer = vbDefault
    Me.Show
    Exit Sub
fine:
 MsgBox err.Description & vbCrLf & g_Settings.dbFile & vbCrLf & sRptFile & vbCrLf & qryFormula, vbCritical
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
 m_Closed = False
End Sub

Private Sub Form_Resize()
On Error GoTo fine
 Me.Move 100, 500, Atap.ScaleWidth - 200, Atap.ScaleHeight - 500
fine:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 m_Closed = True
End Sub

Public Property Get IsClosed() As Boolean
  IsClosed = m_Closed
End Property
Public Property Get Destination() As DestinationConstants
  Destination = m_destination
End Property
