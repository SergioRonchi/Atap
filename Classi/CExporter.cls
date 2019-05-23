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
Public Function esportaXML(Path As String, isNumeroFatture As Boolean, _
                        anno As Integer, daNum As Long, aNum As Long, _
                        Da As Date, A As Date, datePostFix As String) As Boolean
On Error GoTo fine
Dim xml_doc As DOMDocument
Set xml_doc = New DOMDocument

Dim fatture_Node As IXMLDOMElement

Set fatture_Node = xml_doc.createElement("Fatture")
xml_doc.appendChild fatture_Node

InnerEsportaXml xml_doc, False, isNumeroFatture, anno, daNum, aNum, Da, A
InnerEsportaXml xml_doc, True, isNumeroFatture, anno, daNum, aNum, Da, A
xml_doc.Save Path & "/Fatture_" & datePostFix & ".xml"

esportaXML = True

Exit Function
fine:
esportaXML = False
End Function
Private Sub InnerEsportaXml(doc As DOMDocument, isUnep As Boolean, isNumeroFatture As Boolean, _
                        anno As Integer, daNum As Long, aNum As Long, _
                        Da As Date, A As Date)
Dim SQL As String
Dim rs As ADODB.Recordset
Dim Tabella As String
Tabella = IIf(isUnep, "StoricoFattureUnep", "STORICOFATTURE")
 m_prog1 = 0: m_prog2 = 0: RaiseProgress
SQL = "SELECT NumeroFattura, DataFatturaNormale, CodIva, " & _
      "ImportoIVA, CompAdempEuro, CompSfpgEuro, CompNotifEuro, " & _
      "CompDecrIngEuro, " & Tabella & ".NOME, " & Tabella & ".INDIRI, " & Tabella & ".LOCALI, " & Tabella & ".PROV, " & Tabella & ".CAP, " & Tabella & ".PIVA, " & _
      "AnagraficaAvvocati.EMAIL, AnagraficaAvvocati.PEC, AnagraficaAvvocati.STAT, AnagraficaAvvocati.CFISC "
If isUnep Then
  SQL = SQL & ",QUOTA, BIMESTRE"
End If

SQL = SQL & " FROM  AnagraficaAvvocati RIGHT JOIN " & Tabella & " ON AnagraficaAvvocati.CODAVV = " & Tabella & ".CODAVV    "

If isNumeroFatture Then
   SQL = SQL & "WHERE  LEFT(DataFattura,4)='" & anno & "' AND CInt(NumeroFattura)>=" & daNum & " AND CInt(NumeroFattura)<=" & aNum & " " & _
                "ORDER BY Format(NumeroFattura, '00000000')"
Else
   SQL = SQL & "WHERE DataFattura>='" & Format(Da, "YYYYMMDD") & "' AND DataFattura<='" & Format(A, "YYYYMMDD") & "' " & _
                "ORDER BY DataFattura, Format(NumeroFattura, '00000000')"
End If
Set rs = New ADODB.Recordset
rs.Open SQL, g_Settings.DBConnection


While Not rs.EOF

  CreaFatturaXml doc, doc.documentElement, rs, isUnep
  
  m_prog1 = rs.AbsolutePosition * 100 / rs.RecordCount: RaiseProgress
  rs.MoveNext
Wend
End Sub
Private Sub CreaFatturaXml(doc As DOMDocument, parent_node As IXMLDOMElement, rs As ADODB.Recordset, isUnep As Boolean)

Dim nd As IXMLDOMElement
Dim header_node As IXMLDOMElement
Dim target_node As IXMLDOMElement
Dim righe_node As IXMLDOMElement
Dim tot_node As IXMLDOMElement

Dim x As Double
Dim imponibile As Double
Dim imposta As Double
Dim totale As Double


 ' Make the Employee element.
    Set fattura_node = doc.createElement("Fattura")
    parent_node.appendChild fattura_node
    
    Set nd = doc.createElement("Numero")
    nd.appendChild doc.createTextNode(rs("NumeroFattura"))
    fattura_node.appendChild nd
    
    Set nd = doc.createElement("Data")
    nd.appendChild doc.createTextNode(rs("DataFatturaNormale"))
    fattura_node.appendChild nd
    
    Set target_node = doc.createElement("Destinatario")
    fattura_node.appendChild target_node
    
    Set nd = doc.createElement("PartitaIva")
    nd.appendChild doc.createTextNode(rs("PIVA"))
    target_node.appendChild nd
    
    Set nd = doc.createElement("CodiceFiscale")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("CFISC")), "", rs("CFISC")))
    target_node.appendChild nd
    
    Set nd = doc.createElement("PEC")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("PEC")), "", rs("PEC")))
    target_node.appendChild nd
    

    
    Set nd = doc.createElement("Email")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("EMAIL")), "", rs("EMAIL")))
    target_node.appendChild nd
    
    Set nd = doc.createElement("RagioneSociale")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("NOME")), "", rs("NOME")))
    target_node.appendChild nd

    Set nd = doc.createElement("Indirizzo")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("INDIRI")), "", rs("INDIRI")))
    target_node.appendChild nd
    
    Set nd = doc.createElement("Sede")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("LOCALI")), "", rs("LOCALI")))
    target_node.appendChild nd
    
    Set nd = doc.createElement("Provincia")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("PROV")), "", rs("PROV")))
    target_node.appendChild nd
    
    Set nd = doc.createElement("CAP")
    nd.appendChild doc.createTextNode(IIf(IsNull(rs("CAP")), "", rs("CAP")))
    target_node.appendChild nd
    
    Set righe_node = doc.createElement("Righe")
    fattura_node.appendChild righe_node
    
    Set tot_node = doc.createElement("Totale")
    fattura_node.appendChild tot_node
    imponibile = rs("CompNotifEuro") + rs("CompDecrIngEuro") + rs("CompSfpgEuro") + rs("CompAdempEuro")
    imposta = imponibile * rs("ImportoIVA") / 100
    totale = imponibile + imposta
    If isUnep Then
     
      
     CreaRiga doc, "CONTRIBUTO ASSOCIATIVO PROPORZIONALE", righe_node, imponibile, rs("ImportoIVA")
     If InStr(1, rs("BIMESTRE"), "-") > 0 Then
         CreaRiga doc, "CONTRIBUTO ASSOCIATIVO BIMESTRALE " & rs("BIMESTRE"), righe_node, rs("Quota"), 0
      Else
         CreaRiga doc, "CONTRIBUTO ASSOCIATIVO MENSILE " & rs("BIMESTRE"), righe_node, rs("Quota"), 0
     End If
     
       totale = totale + rs("Quota")
     Else
    
    CreaRiga doc, "PER COMPETENZE NOTIFICHE", righe_node, rs("CompNotifEuro"), rs("ImportoIVA")
    CreaRiga doc, "PER COMPETENZE DECRETI INGIUNTIVI", righe_node, rs("CompDecrIngEuro"), rs("ImportoIVA")
    CreaRiga doc, "PER COMPETENZE PIGNORAMENTI", righe_node, rs("CompSfpgEuro"), rs("ImportoIVA")
    CreaRiga doc, "PER COMPETENZE CANCELLERIE", righe_node, rs("CompAdempEuro"), rs("ImportoIVA")
    

    
   End If
   

        Set nd = doc.createElement("Imponibile")
        nd.appendChild doc.createTextNode(Format(imponibile, "#,##0.00"))
        tot_node.appendChild nd
        
        
        Set nd = doc.createElement("Imposta")
        nd.appendChild doc.createTextNode(Format(imposta, "#,##0.00"))
        tot_node.appendChild nd
        
        Set nd = doc.createElement("Totale")
        nd.appendChild doc.createTextNode(Format(totale, "#,##0.00"))
        
        
        tot_node.appendChild nd
   
   
    



End Sub
Private Sub CreaRiga(doc As DOMDocument, descrizione As String, parent_node As IXMLDOMElement, importo As Double, iva As Double)
Dim riga_node As IXMLDOMElement
Dim nd As IXMLDOMElement
Dim x As Double
 If importo > 0 Then
        Set riga_node = doc.createElement("Riga")
        parent_node.appendChild riga_node
        
        Set nd = doc.createElement("Descrizione")
        nd.appendChild doc.createTextNode(descrizione)
        riga_node.appendChild nd
        
        Set nd = doc.createElement("Importo")
        nd.appendChild doc.createTextNode(Format(importo, "#,##0.00"))
        riga_node.appendChild nd
        

        Set nd = doc.createElement("AliquotaIVA")
        nd.appendChild doc.createTextNode(Format(iva, "#,##0.00"))
        riga_node.appendChild nd
        
        x = importo * iva / 100
        Set nd = doc.createElement("Imposta")
        nd.appendChild doc.createTextNode(Format(x, "#,##0.00"))
        riga_node.appendChild nd
        
    End If
End Sub

Public Function Esporta(ByRef oMinmax As MinMax, Path As String, isNumeroFatture As Boolean, _
                        anno As Integer, daNum As Long, aNum As Long, _
                        Da As Date, A As Date, datePostFix As String, fatturePerFile As Long) As Boolean
  On Error GoTo fine
  Dim fattureMsg As String
  Dim iFileNo
  Dim fileName As String
   Dim fileLog As String
  Dim I As Long
  Dim s As String
  Dim Num As String
  Dim listaFatture As Collection
  Dim listaFattureUnep As Collection
  
    fileName = Path & "\FATTURE_" & datePostFix & ".csv"
    fileLog = Path & "\EXPORT_LOG_" & datePostFix & ".txt"
     m_da = Da
     m_a = A
         oMinmax.IntMin = 1000000
         oMinmax.IntMax = 0
         m_prog1 = 0: m_prog2 = 0: RaiseProgress
         
          m_prog1 = 5: m_prog2 = 0: RaiseProgress
         EsportaAnagrafica oMinmax, Path, datePostFix
         
         m_prog1 = 25: m_prog2 = 0: RaiseProgress
         Set listaFatture = EsportaFatture(oMinmax, Path, isNumeroFatture, anno, daNum, aNum, Da, A, True, False, fatturePerFile)
         m_prog1 = 60: m_prog2 = 0: RaiseProgress
         Set listaFattureUnep = EsportaFatture(oMinmax, Path, isNumeroFatture, anno, daNum, aNum, Da, A, True, True, fatturePerFile)
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
        
        If oMinmax.ErrMsg <> "" Then
          Open fileLog For Output As #iFileNo
          Print #iFileNo, oMinmax.ErrMsg
          Close iFileNo
          
          Shell "Notepad " & fileLog, vbMaximizedFocus
        End If
        
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

Private Sub EsportaAnagrafica(ByRef oMinmax As MinMax, Path As String, datePostFix As String)
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
     oMinmax.ErrMsg = oMinmax.ErrMsg & "La cassetta " & rs("CodAvv") & " (" & rs("NOME") & ") ha un codice fiscale errato! Non è stata esportata!" & vbCrLf & "Si prega di correggere prima di esportare di nuovo!" & vbCrLf
    
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
Private Function EsportaFatture(ByRef oMinmax As MinMax, Path As String, isNumeroFatture As Boolean, _
                        anno As Integer, daNum As Long, aNum As Long, _
                        Da As Date, A As Date, addHeader As Boolean, isUnep As Boolean, fatturePerFile As Long) As Collection
Dim rs As ADODB.Recordset
Dim msg As String
Dim row As CExportFatturaRow
Dim Tot As Double
Dim nRiga As Long
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

SQL = SQL & " FROM AnagraficaAvvocati INNER JOIN " & Tabella & " ON AnagraficaAvvocati.CODAVV = " & Tabella & ".CODAVV  "

If isNumeroFatture Then
   SQL = SQL & "WHERE AnagraficaAvvocati.STAT='V' AND LEFT(DataFattura,4)='" & anno & "' AND CInt(NumeroFattura)>=" & daNum & " AND CInt(NumeroFattura)<=" & aNum & " " & _
                "ORDER BY Format(NumeroFattura, '00000000')"
Else
   SQL = SQL & "WHERE AnagraficaAvvocati.STAT='V' AND DataFattura>='" & Format(Da, "YYYYMMDD") & "' AND DataFattura<='" & Format(A, "YYYYMMDD") & "' " & _
                "ORDER BY DataFattura, Format(NumeroFattura, '00000000')"
End If

rs.Open SQL, g_Settings.DBConnection

If addHeader Then msg = CreateIntestazioneFatture & vbCrLf
Dim codPartita As String
Dim numeroRegistrazione As Long
Dim totSenzaIVA As Double
Dim contaFatture As Long
Dim errPIVA As String
Dim scadenza As String
Dim maxNumeroFattura As Long
Dim minNumeroFattura As Long
maxNumeroFattura = oMinmax.IntMax
minNumeroFattura = oMinmax.IntMin
While Not rs.EOF
  If CInt(rs("NumeroFattura")) > maxNumeroFattura Then
     maxNumeroFattura = CInt(rs("NumeroFattura"))
  End If
    If CInt(rs("NumeroFattura")) < minNumeroFattura Then
     minNumeroFattura = CInt(rs("NumeroFattura"))
  End If
  If isUnep Then
    codAliquotaIVA = rs("CodIVA") ' "EC"
    codArticoloIVA = "005" '"160"
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
  Tot = Tot * (1 + rs("ImportoIVA") / 100)
  If isUnep Then
    Tot = Tot + rs("QUOTA")
  End If
  
  
  nRiga = 1
  numeroRegistrazione = rs("NumeroFattura") * 2
  'TESTATA
  row.Init True, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), Tot, 0, codAliquotaIVA, codArticoloIVA, "", g_Settings.CodTestata
  msg = msg & row.CreateRow & vbCrLf
  scadenza = row.Scadenze
  codPartita = row.codPartita
  nRiga = nRiga + 1
  
  If isUnep Then
  
    Set row = New CExportFatturaRow
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), totSenzaIVA, totSenzaIVA * rs("ImportoIVA") / 100, codAliquotaIVA, g_Settings.CodArtIVA, "Quota Proporzionale", Trim(codCE)
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
   
    Set row = New CExportFatturaRow
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), rs("QUOTA"), 0, g_Settings.CodIVAEsente, g_Settings.CodArtIVAEsente, "Quota associativa", Trim(g_Settings.CodQuataFissa)
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  Else
    'UNICA RIGA PER FATTURA TOTALE
    Set row = New CExportFatturaRow
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), totSenzaIVA, totSenzaIVA * rs("ImportoIVA") / 100, codAliquotaIVA, codArticoloIVA, "", Trim(codCE)
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
  End If

  
  'STAMPA INCASSI
    nRiga = 1
    numeroRegistrazione = numeroRegistrazione + 1
    Set row = New CExportFatturaRow
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", "", Tot, 0, "", "", "Incasso da clienti", Trim(g_Settings.CodTestataIncasso)
    row.Causale = "B25"
    row.Valenza = "0"
    row.Fiscale = "0"
    row.ImportoIVA = ""
    row.DataDocumento = row.DataRegistrazione
    msg = msg & row.CreateRow & vbCrLf
    nRiga = nRiga + 1
    Set row = New CExportFatturaRow
    
    row.Init False, CStr(numeroRegistrazione), nRiga, rs("DataFatturaNormale"), rs("NumeroFattura"), "", Trim(rs("PIVA")), -Tot, 0, "", "", "", Trim(g_Settings.CodTestata)
    row.Causale = "B25"
    row.codPartita = codPartita
    row.CFnominativo = Trim(rs("PIVA"))
    row.TipoRiga = "0"
     row.ImportoIVA = ""
    row.Scadenze = scadenza
    msg = msg & row.CreateRow & vbCrLf
    contaFatture = contaFatture + 1
    If contaFatture > fatturePerFile Then
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
oMinmax.IntMax = maxNumeroFattura
oMinmax.IntMin = minNumeroFattura

If errPIVA <> "" Then
  
  oMinmax.ErrMsg = oMinmax.ErrMsg & vbCrLf & "Alcune fatture non hanno la Partita IVA" & vbCrLf & errPIVA
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
