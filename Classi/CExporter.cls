VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event OnDoubleProgress(v1 As Long, v2 As Long)
Private m_prog1 As Long
Private m_prog2 As Long
Private m_da As Date
Private m_a As Date
Public Function Esporta(Path As String, Da As Date, A As Date) As Boolean
  On Error GoTo fine
     
   
     m_da = Da
     m_a = A
  
         m_prog1 = 0: m_prog2 = 0: RaiseProgress
         
          m_prog1 = 5: m_prog2 = 0: RaiseProgress
         EsportaAnagrafica Path
         m_prog1 = 40: m_prog2 = 0: RaiseProgress
         EsportaFatture Path, Da, A
         m_prog1 = 100: m_prog2 = 0: RaiseProgress
         
         Esporta = True
         Exit Function
         
fine:
   MsgBox err.Description, vbOKOnly + vbExclamation
  Esporta = False
End Function
Private Sub RaiseProgress()
 RaiseEvent OnDoubleProgress(m_prog1, m_prog2)
End Sub

Private Sub EsportaAnagrafica(Path As String)
Dim rs As ADODB.Recordset
Dim msg As String
Dim row As CExportAnagraficaRow

Set rs = newAdoRs

sql = "SELECT * FROM ANAGRAFICAAVVOCATI WHERE PIVA IS NOT NULL AND CFISC IS NOT NULL AND CFISC<>'' AND Piva<>''"
      
rs.Open sql, g_Settings.DBConnection


While Not rs.EOF
  Set row = New CExportAnagraficaRow
  row.Init rs("NOME"), rs("CFISC"), rs("PIVA"), rs("INDIRI"), rs("LOCALI"), rs("CAP"), "", rs("PROV"), rs("TELEF"), rs("TELEFCELL"), rs("EMAIL"), rs("FAX"), rs("NOTE1") & " " & rs("NOTE2") & " " & rs("NOTE3")
  msg = msg & row.CreateRow & vbCrLf
  rs.MoveNext
  i = i + 1
  m_prog2 = i * 100 / rs.RecordCount: RaiseProgress
   DoEvents
Wend
Dim iFileNo
Dim fileName As String
iFileNo = FreeFile
fileName = Path & "\ANAGRAFICA_" & Format(Now, "YYYYMMDDHHmm") & ".csv"

Open fileName For Output As #iFileNo
msg = Left(msg, Len(msg) - 2)
Print #iFileNo, msg

Close iFileNo

End Sub
Private Sub EsportaFatture(Path As String, Da As Date, A As Date)
Dim rs As ADODB.Recordset
Dim msg As String
Dim row As CExportFatturaRow
Dim Tot As Double
Dim nRiga As Integer
Set rs = newAdoRs

sql = "SELECT NumeroFattura, DataFatturaNormale,PIVA,CodIva, ImportoIVA, CompAdempEuro, CompSfpgEuro, CompNotifEuro, CompDecrIngEuro FROM STORICOFATTURE " & _
      "WHERE DataFattura>='" & Format(Da, "YYYYMMDD") & "' AND DataFattura<='" & Format(A, "YYYYMMDD") & "' " & _
      "ORDER BY NumeroFattura"
      
rs.Open sql, g_Settings.DBConnection


While Not rs.EOF
  Set row = New CExportFatturaRow
  Tot = rs("CompAdempEuro") + rs("CompSfpgEuro") + rs("CompNotifEuro") + rs("CompDecrIngEuro")
  Tot = Tot * (1 + rs("ImportoIVA") / 100)
  
  row.Init True, rs("NumeroFattura"), 0, rs("DataFatturaNormale"), rs("NumeroFattura"), "", rs("PIVA"), Tot, 0, rs("CodIVA"), ""
  msg = msg & row.CreateRow & vbCrLf
  nRiga = 1
  
  If rs("CompAdempEuro") > 0 Then
    Set row = New CExportFatturaRow
    row.Init False, rs("NumeroFattura"), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", rs("PIVA"), rs("CompAdempEuro"), rs("CompAdempEuro") * rs("ImportoIVA") / 100, rs("CodIVA"), "Adempimenti"
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  End If
  
  If rs("CompSfpgEuro") > 0 Then
    Set row = New CExportFatturaRow
    row.Init False, rs("NumeroFattura"), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", rs("PIVA"), rs("CompSfpgEuro"), rs("CompSfpgEuro") * rs("ImportoIVA") / 100, rs("CodIVA"), "Sfratti"
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  End If
  
  If rs("CompNotifEuro") > 0 Then
    Set row = New CExportFatturaRow
    row.Init False, rs("NumeroFattura"), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", rs("PIVA"), rs("CompNotifEuro"), rs("CompNotifEuro") * rs("ImportoIVA") / 100, rs("CodIVA"), "Notifiche"
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  End If
  
  If rs("CompDecrIngEuro") > 0 Then
    Set row = New CExportFatturaRow
    row.Init False, rs("NumeroFattura"), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", rs("PIVA"), rs("CompDecrIngEuro"), rs("CompDecrIngEuro") * rs("ImportoIVA") / 100, rs("CodIVA"), "Decreti"
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  End If
  
  rs.MoveNext
  i = i + 1
  m_prog2 = i * 100 / rs.RecordCount: RaiseProgress
   DoEvents
Wend
Dim iFileNo
Dim fileName As String
iFileNo = FreeFile
fileName = Path & "\FATTURE_" & Format(Now, "YYYYMMDDHHmm") & ".csv"

Open fileName For Output As #iFileNo
msg = Left(msg, Len(msg) - 2)
Print #iFileNo, msg

Close iFileNo
End Sub
