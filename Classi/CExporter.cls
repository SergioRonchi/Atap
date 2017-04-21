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
Public Function Esporta(Path As String, Da As Date, A As Date, datePostFix As String) As Boolean
  On Error GoTo fine
  Dim fattureMsg As String
  Dim iFileNo
  Dim fileName As String
  Dim I As Long
  Dim s As String
  Dim Num As String
  Dim listaFatture As Collection
  Dim listaFattureUnep As Collection
  
    fileName = Path & "\FATTURE_" & datePostFix & ".csv"
     m_da = Da
     m_a = A
  
         m_prog1 = 0: m_prog2 = 0: RaiseProgress
         
          m_prog1 = 5: m_prog2 = 0: RaiseProgress
         EsportaAnagrafica Path, datePostFix
         
         m_prog1 = 25: m_prog2 = 0: RaiseProgress
         Set listaFatture = EsportaFatture(Path, Da, A, True, False)
         m_prog1 = 60: m_prog2 = 0: RaiseProgress
         Set listaFattureUnep = EsportaFatture(Path, Da, A, True, True)
         m_prog1 = 100: m_prog2 = 0: RaiseProgress

        iFileNo = FreeFile
        
        For I = 1 To listaFatture.count
           Num = I
           s = listaFatture.item(I)
           fileName = Path & "\FATTURE_" & datePostFix & "_" & String(3 - Len(Num), "0") & Num & ".csv"
           Open fileName For Output As #iFileNo
           s = Left(s, Len(s) - 2)
           Print #iFileNo, s
           Close iFileNo
        Next
        
          For I = 1 To listaFattureUnep.count
           Num = I
           s = listaFattureUnep.item(I)
           fileName = Path & "\FATTURE_UNEP_" & datePostFix & "_" & String(3 - Len(Num), "0") & Num & ".csv"
           Open fileName For Output As #iFileNo
           s = Left(s, Len(s) - 2)
           Print #iFileNo, s
           Close iFileNo
        Next
        
'        Open fileName For Output As #iFileNo
'        fattureMsg = Left(fattureMsg, Len(fattureMsg) - 2)
'        Print #iFileNo, fattureMsg
'
'        Close iFileNo
        Esporta = True
         Exit Function
         
fine:
   MsgBox err.Description, vbOKOnly + vbExclamation
  Esporta = False
End Function
Private Sub RaiseProgress()
 RaiseEvent OnDoubleProgress(m_prog1, m_prog2)
End Sub

Private Sub EsportaAnagrafica(Path As String, datePostFix As String)
Dim rs As ADODB.Recordset
Dim msg As String
Dim row As CExportAnagraficaRow

Set rs = newAdoRs

SQL = "SELECT * FROM ANAGRAFICAAVVOCATI WHERE PIVA IS NOT NULL AND CFISC IS NOT NULL AND CFISC<>'' AND Piva<>''"
      
rs.Open SQL, g_Settings.DBConnection

msg = msg & CreateIntestazioneAnagrafica & vbCrLf
While Not rs.EOF
  Set row = New CExportAnagraficaRow
  row.Init rs("NOME"), Trim(UCase(rs("CFISC"))), Trim(rs("PIVA")), rs("INDIRI"), rs("LOCALI"), rs("CAP"), "", rs("PROV"), rs("TELEF"), rs("TELEFCELL"), rs("EMAIL"), rs("FAX"), rs("NOTE1") & " " & rs("NOTE2") & " " & rs("NOTE3")
  If row.DataNascita = "" Or IsDate(row.DataNascita) Then
     msg = msg & row.CreateRow & vbCrLf
    Else
     MsgBox "La cassetta " & rs("CodAvv") & " (" & rs("NOME") & ") ha un codice fiscale errato! Non verrà esportata!" & vbCrLf & "Si prega di correggere prima di esportare di nuovo!", vbOKOnly + vbCritical
   End If
  rs.MoveNext
  I = I + 1
  m_prog2 = I * 100 / rs.RecordCount: RaiseProgress
   DoEvents
Wend
Dim iFileNo
Dim fileName As String
iFileNo = FreeFile
fileName = Path & "\ANAGRAFICA_" & datePostFix & ".csv"

Open fileName For Output As #iFileNo
msg = Left(msg, Len(msg) - 2)
Print #iFileNo, msg

Close iFileNo

End Sub
Private Function EsportaFatture(Path As String, Da As Date, A As Date, addHeader As Boolean, isUnep As Boolean) As Collection
Dim rs As ADODB.Recordset
Dim msg As String
Dim row As CExportFatturaRow
Dim Tot As Double
Dim nRiga As Integer
Set rs = newAdoRs
Dim Tabella As String
Dim fileName As String
Dim codCE As String
Dim codAliquotaIVA As String
Dim codArticoloIVA As String
Dim oListaStringhe As New Collection
Tabella = IIf(isUnep, "StoricoFattureUnep", "STORICOFATTURE")

codCE = g_Settings.CodCompetenze

SQL = "SELECT NumeroFattura, DataFatturaNormale, CodIva, " & _
      "ImportoIVA, CompAdempEuro, CompSfpgEuro, CompNotifEuro, " & _
      "CompDecrIngEuro, AnagraficaAvvocati.PIVA as PIVA_REAL, AnagraficaAvvocati.CFISC as PIVA,AnagraficaAvvocati.CODAVV "
If isUnep Then
  SQL = SQL & ",QUOTA"
  codCE = g_Settings.CodQuotaVariabile
End If

SQL = SQL & " FROM AnagraficaAvvocati INNER JOIN " & Tabella & " ON AnagraficaAvvocati.CODAVV = " & Tabella & ".CODAVV  " & _
      "WHERE AnagraficaAvvocati.STAT='V' AND DataFattura>='" & Format(Da, "YYYYMMDD") & "' AND DataFattura<='" & Format(A, "YYYYMMDD") & "' " & _
      "ORDER BY DataFattura, Format(NumeroFattura, '00000000')"
rs.Open SQL, g_Settings.DBConnection

If addHeader Then msg = CreateIntestazioneFatture & vbCrLf
Dim codPartita As String
Dim numeroRegistrazione As Long
Dim totSenzaIVA As Double
Dim contaFatture As Long
Dim errPIVA As String
Dim scadenza As String
While Not rs.EOF
  If isUnep Then
    codAliquotaIVA = "EC"
    codArticoloIVA = "160"
  Else
    codAliquotaIVA = rs("CodIVA")
    codArticoloIVA = "005"
  End If
  If Trim(rs("PIVA")) = "" Then
    errPIVA = errPIVA & "Fattura senza partita IVA n. " & rs("NumeroFattura") & " CodAvv: " & rs("CodAvV") & vbCrLf
  End If
  Set row = New CExportFatturaRow
  Tot = rs("CompAdempEuro") + rs("CompSfpgEuro") + rs("CompNotifEuro") + rs("CompDecrIngEuro")
  totSenzaIVA = Tot
  If isUnep Then
    Tot = Tot + rs("QUOTA")
  End If
  
  Tot = Tot * (1 + rs("ImportoIVA") / 100)
  nRiga = 1
  numeroRegistrazione = rs("NumeroFattura") * 2
  'TESTATA
  row.Init True, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), Tot, 0, codAliquotaIVA, codArticoloIVA, "", g_Settings.CodTestata
  msg = msg & row.CreateRow & vbCrLf
  scadenza = row.Scadenze
  codPartita = row.codPartita
  nRiga = nRiga + 1
  
  'UNICA RIGA PER FATTURA TOTALE
    Set row = New CExportFatturaRow
    'row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), totSenzaIVA, totSenzaIVA * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Importo Competenze " & g_Settings.Sede, codCE
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), totSenzaIVA, totSenzaIVA * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "", codCE
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
    
'RIGHE SEPARATE
'  If rs("CompAdempEuro") > 0 Then
'    Set row = New CExportFatturaRow
'    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("CompAdempEuro"), rs("CompAdempEuro") * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Adempimenti", codCE
'    msg = msg & row.CreateRow & vbCrLf
'    nRiga = nRiga + 1
'  End If
'
'  If rs("CompSfpgEuro") > 0 Then
'    Set row = New CExportFatturaRow
'    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("CompSfpgEuro"), rs("CompSfpgEuro") * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Sfratti", codCE
'    msg = msg & row.CreateRow & vbCrLf
'    nRiga = nRiga + 1
'  End If
'
'  If rs("CompNotifEuro") > 0 Then
'    Set row = New CExportFatturaRow
'    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("CompNotifEuro"), rs("CompNotifEuro") * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Notifiche", codCE
'    msg = msg & row.CreateRow & vbCrLf
'    nRiga = nRiga + 1
'  End If
'
'  If rs("CompDecrIngEuro") > 0 Then
'    Set row = New CExportFatturaRow
'    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("CompDecrIngEuro"), rs("CompDecrIngEuro") * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Decreti", codCE
'    msg = msg & row.CreateRow & vbCrLf
'    nRiga = nRiga + 1
'  End If
  
   If isUnep Then
    If rs("QUOTA") > 0 Then
        Set row = New CExportFatturaRow
        row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("QUOTA"), rs("QUOTA") * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "Quota associativa", g_Settings.CodQuataFissa
        msg = msg & row.CreateRow & vbCrLf
        nRiga = nRiga + 1
    End If
  End If
  
  'STAMPA INCASSI
    nRiga = 1
    numeroRegistrazione = numeroRegistrazione + 1
    Set row = New CExportFatturaRow
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", "", Tot, 0, "", "", "Incasso da clienti", g_Settings.CodTestataIncasso
    row.Causale = "B25"
    row.Valenza = "0"
    row.Fiscale = "0"
    row.ImportoIVA = ""
    row.DataDocumento = row.DataRegistrazione
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
    Set row = New CExportFatturaRow
    
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), -Tot, 0, "", "", "", g_Settings.CodTestata
    row.Causale = "B25"
    row.codPartita = codPartita
    row.CFnominativo = Trim(rs("PIVA"))
    row.TipoRiga = "0"
     row.ImportoIVA = ""
    row.Scadenze = scadenza
    msg = msg & row.CreateRow & vbCrLf
    contaFatture = contaFatture + 1
    If contaFatture > 100 Then
      oListaStringhe.Add msg
      contaFatture = 0
      If addHeader Then msg = CreateIntestazioneFatture & vbCrLf
    End If
  rs.MoveNext
  
  I = I + 1
  m_prog2 = I * 100 / rs.RecordCount: RaiseProgress
   DoEvents
Wend

If contaFatture > 0 Then
   oListaStringhe.Add msg
   
End If

If errPIVA <> "" Then
  MsgBox "Alcune fatture non hanno la Partita IVA" & vbCrLf & errPIVA
End If
Set EsportaFatture = oListaStringhe
End Function
Private Function CreateIntestazioneAnagrafica() As String
Dim s As String
      s = s & "codice;"
 
    s = s & "CodiceFiscale;"
    s = s & "FLCliente;"
    s = s & "FLCAF;"
    s = s & "FLIntermediario;"
    s = s & "IDTipologia;"
    s = s & "COGNOME;"
    s = s & "nome;"
    s = s & "Denominazione;"
    s = s & "Sesso;"
    s = s & "DataNascita;"
    s = s & "IDStatoCivile;"
    s = s & "ComuneNascita;"
    s = s & "IDProvinciaNascita;"
    s = s & "CodComuneNascita;"
    s = s & "PartitaIVA;"
    s = s & "Sigla;"
    s = s & "CodiceNaturaGiuridica;"
    
    s = s & "IndirizzoResidenza;"
    s = s & "NumeroCivicoResidenza;"
    s = s & "ComuneResidenza;"
    s = s & "IDProvinciaResidenza;"
    s = s & "CodComuneResidenza;"
    s = s & "CAPResidenza;"
    s = s & "FLPrincipale;"
    
    s = s & "NazioneEstera;"
    s = s & "CodCatastNazione;"
    s = s & "CodEsteroNazione;"
    s = s & "StatoFederaleEstero;"
    s = s & "LocalitaEstero;"
    s = s & "IndirizzoEstero;"
    s = s & "NumeroCivicoEstero;"
    s = s & "ZIP;"
    s = s & "FLNazionalitaEstera;"
    s = s & "FLCittadinanzaItalia;"
    s = s & "CFEstero;"
    s = s & "FLPrincipaleEstero;"
    s = s & "PartitaIVAEstera;"
    s = s & "IDISO;"
    s = s & "CFStabile;"
    
    s = s & "IndirizzoDomFisc;"
    s = s & "NumeroCivico;"
    s = s & "ComuneDomFisc;"
    s = s & "IDProvinciaDomFisc;"
    s = s & "CodComuneDomFisc;"
    s = s & "CAPDomFisc;"
    s = s & "Domfiscale_indirizzo_primario;"
    
    s = s & "telefono;"
    s = s & "fax;"
    s = s & "cellulare;"
    s = s & "Indirizzo_email;"
    s = s & "Indirizzo_URL;"
    
    s = s & "Data_inizio_rapporto;"
    s = s & "Note;"
    s = s & "CodiceFiscaleAzienda;"
    s = s & "FLCliFor"
    
    CreateIntestazioneAnagrafica = s
End Function
Private Function CreateIntestazioneFatture() As String
Dim s As String
    s = s & "N_mov;"
    s = s & "N_riga;"
    s = s & "Causale;"
    s = s & "Valenza;"
    s = s & "Fiscale;"
    s = s & "DataRegistrazione;"
    s = s & "DataCompetenzaBil;"
    s = s & "DataCompetenzaIVA;"
    s = s & "DataDocumento;"
    s = s & "NumDocumento;"
    s = s & "DataDocRiferimento;"
    s = s & "DescrizioneMov;"
    s = s & "CodProvvisorio;"
    s = s & "DocIVAEsigibilDiff;"
    s = s & "RegimeIVAEsigibilDiff;"
    s = s & "TipoPagamento;"
    s = s & "ModPagamento;"
    s = s & "IncassoAccreditoDifferito;"
    s = s & "CodPartita;"
    s = s & "Scadenze;"
    s = s & "NuovaScadenza;"
    s = s & "RegistroIVA;"
    s = s & "ProtocolloIVA;"
    s = s & "ProtocolloArrivo;"
    s = s & "RegIVAVendite;"
    s = s & "ImportoRitenuta;"
    s = s & "ImportoEnasarco;"
    s = s & "CodiceBarre;"
    s = s & "Allegato;"
    s = s & "CodProvenienza;"
    s = s & "CodiceFiscaleFunebri;"
    s = s & "CognomeNomeFunebri;"
    
    s = s & "TipoRiga;"
    s = s & "CodConto;"
    s = s & "CodNominativo;"
    s = s & "CFnominativo;"
    s = s & "DescrizioneRiga;"
    s = s & "ImportoRiga;"
    s = s & "ImportoIVA;"
    s = s & "ImportoValuta;"
    s = s & "CodValutaEstera;"
    s = s & "Cambio;"
    s = s & "CodAliquotaIVA;"
    s = s & "CodArticoloIVA;"
    s = s & "CodCentroCosto;"
    s = s & "CodAttISTAT;"
    s = s & "NumUnitProd;"
    s = s & "TipologiaSpesaSanitaria;"
    s = s & "TipoRettifica;"
    s = s & "ContoRettifica;"
    s = s & "DataCompBilancioRett;"
    s = s & "PeriodoRettDa;"
    s = s & "PeriodoRettA;"

    CreateIntestazioneFatture = s
End Function
