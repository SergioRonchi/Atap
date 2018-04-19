VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub UpdateVeryOld(CN As ADODB.Connection)
If Not ExistMDBTable(CN, "LockTable") Then
     CN.Execute "CREATE TABLE LockTable (TabID TEXT(50) PRIMARY KEY, Locked TEXT(50) DEFAULT 'NONE')"
     CN.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('ADEMPI','NONE')"
     CN.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('AnagraficaAvvocati','NONE')"
     CN.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('DecretiIngiuntivi','NONE')"
     CN.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('Notifiche','NONE')"
     CN.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('SFRATTI','NONE')"
End If
If Not ExistMDBTable(CN, "NumeriFattura") Then
     CN.Execute "CREATE TABLE NumeriFattura (Numero LONG)"
     CN.Execute "INSERT INTO NumeriFattura (Numero) VALUES (1)"
End If
If Not ExistMDBField(CN, "Anticipi", "Locked") Then
     CN.Execute "ALTER TABLE Anticipi ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE Anticipi SET Locked='NONE'"
 End If

If Not ExistMDBField(CN, "ADEMPI", "Locked") Then
     CN.Execute "ALTER TABLE ADEMPI ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE ADEMPI SET Locked='NONE'"
     CN.Execute "ALTER TABLE AnagraficaAvvocati ADD COLUMN Locked TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE  AnagraficaAvvocati SET Locked='NONE'"
     CN.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE  DecretiIngiuntivi SET Locked='NONE'"
     CN.Execute "ALTER TABLE Notifiche ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE  Notifiche SET Locked='NONE'"
     CN.Execute "ALTER TABLE SFRATTI ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     CN.Execute "UPDATE  SFRATTI SET Locked='NONE'"

End If
'Rivedi ordinamento avvocati
If Not ExistMDBField(CN, "Parametri", "XXX") Then
   'Cerca dupplicati
   sql = "SELECT Anticipi.Descrizione, Anticipi.CodiceAnticipi, Anticipi_1.CodiceAnticipi " & _
       "FROM Anticipi, Anticipi AS Anticipi_1 " & _
       "WHERE (((Anticipi.Descrizione)=[Anticipi_1].[Descrizione]) AND ((Anticipi.CodiceAnticipi)>[Anticipi_1].[CodiceAnticipi]));"
   Set rsAv = newAdoRs
   rsAv.Open sql, CN
   While Not rsAv.EOF
     Code = rsAv(1)
     Descr = rsAv(0) & " " & Code
     CN.Execute "UPDATE ANTICIPI SET DESCRIZIONE='" & Replace(Descr, "'", "''") & "' WHERE CodiceAnticipi='" & Replace(Code, "'", "''") & "'; "
     rsAv.MoveNext
   Wend
   rsAv.Close
  CN.Execute "CREATE UNIQUE INDEX uniDescr ON ANTICIPI (Descrizione)"
  CN.Execute "ALTER TABLE Parametri ADD COLUMN XXX logical"
  Set rsAv = GetADORecordset("AnagraficaAvvocati", "CodAvv", "1=1 ORDER BY NumOrdinamento", CN)
  OpenProgress ("Attendere... Aggiornamento in corso!")
  While Not rsAv.EOF
   UpdateProgress rsAv.AbsolutePosition * 100 / rsAv.RecordCount
   Ord = Ord + 50
   codAvv = rsAv(0)
   CN.Execute "UPDATE AnagraficaAvvocati set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE ADEMPI set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE DecretiIngiuntivi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE Notifiche set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtAssegniCircolari set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtEstrattoConto set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtGiornalieraAdempimenti set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtGiornalieraDecretiIngiuntivi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtGiornalieraNotifiche set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtGiornalieraSfrattiPig set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE PrtSaldi set NumOrdinamento=" & Ord & " WHERE Codice='" & codAvv & "'"
   CN.Execute "UPDATE PrtSospesi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE Saldi set NumOrdinamento=" & Ord & " WHERE Codice='" & codAvv & "'"
   CN.Execute "UPDATE Sfratti set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   CN.Execute "UPDATE StoricoFatture set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   rsAv.MoveNext
  Wend
  CloseProgress


End If
CN.Execute "UPDATE ANAGRAFICAAVVOCATI SET CASSETTAROTTA='N' WHERE CASSETTAROTTA Is Null"
CN.Execute "UPDATE ANAGRAFICAAVVOCATI SET STAT='V' WHERE STAT Is Null"
If Not ExistMDBField(CN, "prtEstrattoConto", "CodAutorita") Then
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN CodAutorita VARCHAR(2)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Esenzione VARCHAR(1)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN FormulaEsec VARCHAR(1)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN NumeroIngiunzione VARCHAR(6)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN NumeroRuolo VARCHAR(7)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Localita1 VARCHAR(35)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Nota VARCHAR(14)"
  CN.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Descrizione VARCHAR(50)"
  CN.Execute "ALTER TABLE prtSospesi ADD COLUMN Nota VARCHAR(14)"
End If
If Not ExistMDBField(CN, "DecretiIngiuntivi", "IDCod") Then
  CN.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN IDCod COUNTER"
  CN.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN Ricorrente Text(35)"
  CN.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN Debitore Text(35)"
  CN.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN DesrSpese Text(35)"
End If
If Not ExistMDBField(CN, "Parametri", "IVA") Then
  CN.Execute "ALTER TABLE Parametri ADD COLUMN IVA DOUBLE"
  CN.Execute "UPDATE Parametri SET IVA=20.0"
End If
If Not ExistMDBField(CN, "AnagraficaAvvocati", "ID") Then
  CN.Execute "ALTER TABLE AnagraficaAvvocati ADD COLUMN ID COUNTER"
  CN.Execute "ALTER TABLE AnagraficaAvvocati ALTER COLUMN NOME Text(50)"
End If
If Not ExistMDBField(CN, "ADEMPI", "SIGLA") Then
  CN.Execute "ALTER TABLE ADEMPI ADD COLUMN SIGLA VARCHAR(15)"
  CN.Execute "ALTER TABLE Sfratti ADD COLUMN SIGLA VARCHAR(15)"
  CN.Execute "ALTER TABLE Notifiche ADD COLUMN SIGLA VARCHAR(15)"
  CN.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN SIGLA VARCHAR(15)"
  CN.Execute "ALTER TABLE ADEMPI ADD COLUMN SIGLACH VARCHAR(15)"
  CN.Execute "ALTER TABLE Sfratti ADD COLUMN SIGLACH VARCHAR(15)"
  CN.Execute "ALTER TABLE Notifiche ADD COLUMN SIGLACH VARCHAR(15)"
  CN.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN SIGLACH VARCHAR(15)"

  CN.Execute "UPDATE DecretiIngiuntivi SET SIGLACH='OLD'"
  CN.Execute "UPDATE ADEMPI SET SIGLACH='OLD'"
  CN.Execute "UPDATE Sfratti SET SIGLACH='OLD'"
  CN.Execute "UPDATE Notifiche SET SIGLACH='OLD'"

  CN.Execute "UPDATE DecretiIngiuntivi SET SIGLA='OLD'"
  CN.Execute "UPDATE ADEMPI SET SIGLA='OLD'"
  CN.Execute "UPDATE Sfratti SET SIGLA='OLD'"
  CN.Execute "UPDATE Notifiche SET SIGLA='OLD'"

End If
If Not ExistMDBField(CN, "Notifiche", "IDCod") Then
  CN.Execute "ALTER TABLE Notifiche ADD COLUMN IDCod COUNTER"
  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN NaturaAtto Text(35)"
  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte1 Text(35)"
  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte2 Text(35)"
  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN DesrSpese Text(35)"
  CN.Execute "ALTER TABLE Sfratti ALTER COLUMN Localita1 Text(40)"

End If
If Not ExistMDBField(CN, "PrtSospesi", "DATA_EVASIONE") Then
  CN.Execute "ALTER TABLE PrtSospesi ADD DATA_EVASIONE Date"

  CN.Execute "ALTER TABLE PrtSospesi ADD SALDO_PRECEDENTE Double"
  CN.Execute "ALTER TABLE PrtSospesi ADD NUM_EST_CONTO Long"
  CN.Execute "ALTER TABLE PrtSospesi ADD DATA_EST_CONTO DAte"
  CN.Execute "ALTER TABLE PrtSospesi ADD PROVVISORIO Text(1)"
  CN.Execute "ALTER TABLE PrtSospesi ADD DATA_INIZIO Text(10)"
  CN.Execute "ALTER TABLE PrtSospesi ADD DATA_FINE Text(10)"

  CN.Execute "ALTER TABLE PrtSospesi ADD QtaMarche Long"
  CN.Execute "ALTER TABLE PrtSospesi ADD QtaFotocopie Long"
  CN.Execute "ALTER TABLE PrtSospesi ADD QtaCopie Text(5)"
  CN.Execute "ALTER TABLE PrtSospesi ADD QtaDirittiCancelleria Long"

  End If

If Not ExistMDBField(CN, "Sfratti", "IDCod") Then
  CN.Execute "ALTER TABLE Sfratti ADD COLUMN IDCod COUNTER"
  CN.Execute "ALTER TABLE Sfratti ALTER COLUMN Localita1 Text(40)"
  CN.Execute "ALTER TABLE Sfratti ALTER COLUMN Parte1 Text(35)"
  CN.Execute "ALTER TABLE Sfratti ALTER COLUMN Parte2 Text(35)"
  CN.Execute "ALTER TABLE Sfratti ALTER COLUMN DesrSpese Text(35)"
End If
If Not ExistMDBField(CN, "ADEMPI", "IDCod") Then
  CN.Execute "ALTER TABLE ADEMPI ADD COLUMN IDCod COUNTER"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese1 Text(35)"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese2 Text(35)"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese3 Text(35)"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese4 Text(35)"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese5 Text(35)"
  CN.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese6 Text(35)"

  CN.Execute "ALTER TABLE PrtAssegniCircolari ALTER COLUMN LOCALI Text(40)"
  CN.Execute "ALTER TABLE PrtAssegniCircolari ALTER COLUMN Nome Text(50)"

  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN LOCALI Text(40)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Nome Text(50)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Parte1 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Parte2 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE1 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE2 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE3 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE4 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE5 Text(35)"
  CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE6 Text(35)"

  CN.Execute "ALTER TABLE PrtFattProv ALTER COLUMN LOCALI Text(40)"
  CN.Execute "ALTER TABLE PrtFattProv ALTER COLUMN Nome Text(50)"

  CN.Execute "ALTER TABLE PrtGiornalieraDecretiIngiuntivi ALTER COLUMN Ricorrente Text(35)"
  CN.Execute "ALTER TABLE PrtGiornalieraDecretiIngiuntivi ALTER COLUMN Debitore Text(35)"

  CN.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN parte1 Text(35)"
  CN.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN parte2 Text(35)"
  CN.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN Localita1 Text(40)"

  CN.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN parte1 Text(35)"
  CN.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN parte2 Text(35)"
  CN.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN Localita1 Text(40)"

  CN.Execute "ALTER TABLE PrtSaldi ALTER COLUMN Nome Text(50)"

  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN Nome Text(50)"


  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN LOCALI Text(40)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN PARTE1 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN PARTE2 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE1 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE2 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE3 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE4 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE5 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE6 Text(35)"
  CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN Localita1 Text(35)"

'Non rimuovo per il momento le tabelle
'  cn.Execute "DROP TABLE Filter"
'  cn.Execute "DROP TABLE LockAdempimenti"
'  cn.Execute "DROP TABLE LockDecreti"
'  cn.Execute "DROP TABLE LockNotifiche"
'  cn.Execute "DROP TABLE LockPignoramenti"

End If

If Not ExistMDBTable(CN, "LockPrt") Then
  CN.Execute "CREATE TABLE LockPrt (" & _
                   "PrtAssegniCircolari Logical," & _
                   "PrtEstrattoConto Logical," & _
                   "PrtFattProv Logical," & _
                   "PrtGiornalieraAdempimenti Logical," & _
                   "PrtGiornalieraDecretiIngiuntivi Logical," & _
                   "PrtGiornalieraNotifiche Logical," & _
                   "PrtGiornalieraSfrattiPig Logical," & _
                   "PrtSaldi Logical," & _
                   "PrtSospesi Logical" & _
                   ");"
  CN.Execute "INSERT INTO LockPrt(PrtAssegniCircolari," & _
                   "PrtEstrattoConto ," & _
                   "PrtFattProv ," & _
                   "PrtGiornalieraAdempimenti ," & _
                   "PrtGiornalieraDecretiIngiuntivi ," & _
                   "PrtGiornalieraNotifiche ," & _
                   "PrtGiornalieraSfrattiPig ," & _
                   "PrtSaldi ," & _
                   "PrtSospesi) VALUES " & _
                   "(False,False,False,False,False,False,False,False,False);"
End If
If Not ExistMDBField(CN, "PrtEStrattoConto", "DataRegistrazione") Then
   CN.Execute "ALTER TABLE PrtEstrattoConto ADD COLUMN DataRegistrazione Text(10) NOT NULL"
   CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DATA_EVASIONE Text(10) NOT NULL"
   CN.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DATARegistrazione Text(10) NOT NULL"
   CN.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DATA_EVASIONE Text(10) NOT NULL"
End If
If Not ExistMDBTable(CN, "TempSaldi") Then
     CN.Execute "SELECT * INTO TempSaldi FROM PrtAssegniCircolari "
     CN.Execute "DELETE * FROM TempSaldi"
End If

If Not ExistMDBField(CN, "Parametri", "CECodTestata") Then
     CN.Execute "ALTER TABLE Parametri ADD CECodTestata Text(12)"
     CN.Execute "ALTER TABLE Parametri ADD CECodQVar Text(12)"
     CN.Execute "ALTER TABLE Parametri ADD CECodQFixed Text(12)"
     CN.Execute "ALTER TABLE Parametri ADD CECodCompetenze  Text(12)"

     CN.Execute "UPDATE Parametri SET CECodTestata='100301050', CECodQVar='600151051', CECodQFixed='600151050', CECodCompetenze='600151052'"
End If

'Elimina duplicati nei saldi
Dim eliminati As String
Dim key As String

Set rs = New ADODB.Recordset
rs.Open "SELECT Saldi.Codice, Saldi.Chiusura, Saldi.SaldoAdemp, Saldi.SaldoAdempEuro, Saldi.SaldoSfpg, Saldi.SaldoSfpgEuro, Saldi.SaldoNotif, Saldi.SaldoNotifEuro, Saldi.SaldoDecrIng, Saldi.SaldoDecrIngEuro, Saldi.Stato, Saldi.SaldoTotale, Saldi.SaldoTotaleEuro, Saldi.PROG_Saldi, Saldi.Commento, Saldi.NumOrdinamento " & _
        " FROM Saldi INNER JOIN Saldi AS Saldi_1 ON Saldi.Codice = Saldi_1.Codice" & _
        " WHERE (((Saldi.CHIUSURA)<>[Saldi_1].[cHIUSURA])) ORDER BY  Saldi.CHIUSURA DESC;", CN
While Not rs.EOF
  key = "$" & rs!codice & "$"
  If InStr(1, eliminati, key) = 0 Then
    CN.Execute "DELETE * FROM Saldi WHERE Codice='" & rs!codice & "' AND Chiusura<>'" & rs!Chiusura & "'"
    eliminati = eliminati + key
  End If

  rs.MoveNext

Wend

' UNEP
'-----------------------------------------------------------------------------------------------------------------------------------------
If Not ExistMDBField(CN, "Sfratti", "Crono") Then
  CN.Execute "ALTER TABLE Sfratti ADD COLUMN IsUNEP Logical"
  CN.Execute "UPDATE Sfratti SET  IsUNEP=false"

  CN.Execute "ALTER TABLE Sfratti ADD COLUMN Crono TEXT(50)"

End If

If Not ExistMDBField(CN, "Notifiche", "Crono") Then
  CN.Execute "ALTER TABLE Notifiche ADD COLUMN IsUNEP Logical"
  CN.Execute "UPDATE Notifiche SET  IsUNEP=false"

  CN.Execute "ALTER TABLE Notifiche ADD COLUMN Crono TEXT(50)"

End If

If Not ExistMDBField(CN, "Parametri", "QuotaSoci") Then
  CN.Execute "ALTER TABLE Parametri ADD COLUMN QuotaSoci Double"
  CN.Execute "UPDATE Parametri SET  QuotaSoci=55"

  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte1 VARCHAR(35) NULL"
  CN.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte2 VARCHAR(35) NULL"

  CN.Execute "CREATE TABLE PrtGiornalieraNotificheUNEP (Crono TEXT(50), DescrizioneTribunale TEXT(30), CODAVV TEXT(10), NumeroAtto LONG, DataRegistrazione TEXT(10), " & _
             "parte1 TEXT(35), parte2 TEXT(35), [Note] TEXT(14), Localita1 TEXT(40), Descrizione TEXT(50), ImpSaldo DOUBLE, Valuta TEXT(4), NumOrdinamento LONG, " & _
             "DataPresentazione TEXT(10), DataRestituzione TEXT(10), DataNotifica TEXT(10), DATA_INIZIO TEXT(10), DATA_FINE TEXT(10))"

  CN.Execute "CREATE TABLE PrtGiornalieraSfrattiPigUNEP (Crono TEXT(50), DescrizioneTribunale TEXT(30), CODAVV TEXT(10), NumeroAtto LONG, DataRegistrazione TEXT(10), DataPresentazione TEXT(10)," & _
             "parte1 TEXT(35), parte2 TEXT(35), Localita1 TEXT(40), Descrizione TEXT(50), ImpSaldo DOUBLE, Valuta TEXT(4), NumOrdinamento LONG, " & _
              "DataRestituzione TEXT(10),  DATA_INIZIO TEXT(10), DATA_FINE TEXT(10))"

  CN.Execute "ALTER TABLE LockPrt ADD PrtGiornalieraNotificheUNEP Logical, PrtGiornalieraSfrattiPigUNEP Logical"
  CN.Execute "UPDATE LockPrt SET PrtGiornalieraNotificheUNEP=false, PrtGiornalieraSfrattiPigUNEP=false"
End If

Set rs = New ADODB.Recordset
rs.Open "SELECT CodiceTribunale FROM TribunaliAppartenenza WHERE CodiceTribunale='UNEP'", CN

If rs.EOF Then
  CN.Execute "INSERT INTO TribunaliAppartenenza (CodiceTribunale, DescrizioneTribunale) VALUES('UNEP', 'MILANO UNEP')"
End If

If Not ExistMDBTable(CN, "PrtEstrattoContoUNEP") Then
  CN.Execute "SELECT * INTO PrtEstrattoContoUNEP FROM PrtEstrattoConto"

End If

If Not ExistMDBTable(CN, "TempSaldiUNEP") Then
  CN.Execute "SELECT * INTO TempSaldiUNEP FROM TempSaldi"

End If

If Not ExistMDBTable(CN, "PrtAssegniCircolariUNEP") Then
  CN.Execute "SELECT * INTO PrtAssegniCircolariUNEP FROM PrtAssegniCircolari"

End If

If Not ExistMDBTable(CN, "StoricoFattureUNEP") Then
  CN.Execute "SELECT * INTO StoricoFattureUNEP FROM StoricoFatture"
  CN.Execute "DELETE * FROM StoricoFattureUNEP"
End If


If Not ExistMDBTable(CN, "SaldiUNEP") Then

  CN.Execute "SELECT * INTO SaldiUNEP FROM Saldi"
  CN.Execute "UPDATE SaldiUNEP SET SaldoAdemp=0, SaldoAdempEuro=0,  SaldoSfpg=0, SaldoSfpgEuro=0, SaldoNotif=0, SaldoNotifEuro=0, SaldoDecrIng=0, SaldoDecrIngEuro=0, SaldoTotale=0, SaldoTotaleEuro=0, Commento=''"
End If

If Not ExistMDBTable(CN, "PrtSaldiUNEP") Then
  CN.Execute "SELECT * INTO PrtSaldiUNEP FROM PrtSaldi"
  CN.Execute "DELETE * FROM PrtSaldiUNEP"
End If

If Not ExistMDBTable(CN, "PrtSospesiUNEP") Then
  CN.Execute "SELECT * INTO PrtSospesiUNEP FROM PrtSospesi"
  CN.Execute "DELETE * FROM PrtSospesiUNEP"
End If


If Not ExistMDBField(CN, "LockPrt", "PrtAssegniCircolariUNEP") Then
  CN.Execute "ALTER TABLE  LockPrt ADD PrtAssegniCircolariUNEP Logical, PrtEstrattoContoUNEP Logical"

  CN.Execute "UPDATE LockPrt SET PrtAssegniCircolariUNEP=false, PrtEstrattoContoUNEP=false "
End If

If Not ExistMDBField(CN, "PrtEstrattoContoUNEP", "DATA_PRESENTAZIONE") Then

  CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD Crono Text(50)"


  CN.Execute "ALTER TABLE  PrtEstrattoConto ADD Crono Text(50)"


  CN.Execute "ALTER TABLE  prtSospesi ADD Crono Text(50)"

  CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"


  CN.Execute "ALTER TABLE  PrtEstrattoConto ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"


  CN.Execute "ALTER TABLE  prtSospesi ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"

End If

If Not ExistMDBTable(CN, "PrtData") Then
   CN.Execute "CREATE TABLE PrtData (Bimestre Integer, BimestreAnno Integer)"

End If


If Not ExistMDBField(CN, "LockPrt", "PrtSaldiUNEP") Then
  CN.Execute "ALTER TABLE  LockPrt ADD PrtSaldiUNEP Logical, PrtSospesiUNEP Logical"

  CN.Execute "UPDATE LockPrt SET PrtSaldiUNEP=false, PrtSospesiUNEP=false "
End If

If Not ExistMDBField(CN, "Date_EstrattiConto", "DATA_ULTIMO_ESTCONTO_UNEP") Then
  CN.Execute "ALTER TABLE  Date_EstrattiConto ADD DATA_ULTIMO_ESTCONTO_UNEP TEXT(10), DATA_FATTURA_UNEP TEXT(10)"

End If

If Not ExistMDBField(CN, "StoricoFattureUNEP", "Bimestre") Then
  CN.Execute "ALTER TABLE  StoricoFattureUNEP ADD Bimestre TEXT(50), QUOTA DOUBLE"

End If
'-----------------------------------------------------------------------------------------------

If Not ExistMDBTable(CN, "SFRATTI_UNEP") Then
  CN.Execute "SELECT * INTO SFRATTI_UNEP FROM SFRATTI"
  CN.Execute "DELETE * FROM SFRATTI_UNEP WHERE IsUNEP=False"
  CN.Execute "DELETE * FROM SFRATTI WHERE IsUNEP=True"

  CN.Execute "ALTER TABLE  SFRATTI DROP COLUMN IsUNEP"
  CN.Execute "ALTER TABLE  SFRATTI_UNEP DROP COLUMN IsUNEP"
  CN.Execute "INSERT INTO LockTable(TabId) VALUES('SFRATTI_UNEP')"
End If

If Not ExistMDBTable(CN, "NOTIFICHE_UNEP") Then
  CN.Execute "SELECT * INTO NOTIFICHE_UNEP FROM NOTIFICHE"
  CN.Execute "DELETE * FROM NOTIFICHE_UNEP WHERE IsUNEP=False"
  CN.Execute "DELETE * FROM NOTIFICHE WHERE IsUNEP=True"

  CN.Execute "ALTER TABLE  NOTIFICHE DROP COLUMN IsUNEP"
  CN.Execute "ALTER TABLE  NOTIFICHE_UNEP DROP COLUMN IsUNEP"
  CN.Execute "INSERT INTO LockTable(TabId) VALUES('NOTIFICHE_UNEP')"

End If

If Not ExistMDBField(CN, "PrtData", "Tipo") Then
  CN.Execute "ALTER TABLE  PrtData ADD Tipo Integer"

End If

If Not ExistMDBField(CN, "Parametri", "LimiteSaldo") Then
  CN.Execute "ALTER TABLE  Parametri ADD LimiteSaldo DOUBLE"
  CN.Execute "UPDATE Parametri SET LimiteSaldo=5.16456"

End If
If Not ExistMDBField(CN, "Parametri", "Banca") Then
  CN.Execute "ALTER TABLE  Parametri ADD Banca TEXT(200), IBAN TEXT(50)"
  CN.Execute "UPDATE Parametri SET Banca='BNL AG. 4373 ', IBAN='IT 45 J 01005 01773 000000220006'"

End If

If Not ExistMDBField(CN, "Parametri", "Telefono") Then
  CN.Execute "ALTER TABLE  Parametri ADD COLUMN Telefono TEXT(20)"
  CN.Execute "ALTER TABLE  Sfratti ALTER COLUMN CodicePignoramenti TEXT(5)"
  CN.Execute "ALTER TABLE  Sfratti_UNEP ALTER COLUMN CodicePignoramenti TEXT(5)"

  CN.Execute "ALTER TABLE Sfratti_Unep ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,DataRegistrazione,NumeroAtto,CodTribunaleApp)"
  CN.Execute "ALTER TABLE StoricoFattureUNEP ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,NumeroFattura,DataFattura)"
  CN.Execute "ALTER TABLE Notifiche_UNEP ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,DataRegistrazione,NumeroAtto,CodTribunaleApp)"

End If

If Not ExistMDBField(CN, "Parametri", "CodIVA") Then
  CN.Execute "ALTER TABLE  Parametri ADD COLUMN CodIVA TEXT(5)"
  CN.Execute "UPDATE Parametri SET  CodIVA='022'"


  CN.Execute "ALTER TABLE  StoricoFatture ADD COLUMN CodIVA TEXT(5)"
  CN.Execute "UPDATE StoricoFatture SET  CodIVA='022' WHERE DataFattura>='20131001'"
  CN.Execute "UPDATE StoricoFatture SET  CodIVA='020' WHERE DataFattura<'20131001'"
  CN.Execute "UPDATE StoricoFatture SET  ImportoIVA=22 WHERE DataFattura>='20131001'"

  CN.Execute "ALTER TABLE  StoricoFattureUNEP ADD COLUMN CodIVA TEXT(5)"
  CN.Execute "UPDATE StoricoFattureUNEP SET  CodIVA='022' WHERE DataFattura>='20131001'"

End If


If Not ExistMDBField(CN, "PrtEstrattoContoUNEP", "Quota") Then
   CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD COLUMN Quota Double"
   CN.Execute "ALTER TABLE  TempSaldiUNEP ADD COLUMN Quota Double"
   CN.Execute "ALTER TABLE  PrtAssegniCircolariUNEP ADD COLUMN Quota Double"
   CN.Execute "ALTER TABLE  PrtSospesiUNEP ADD COLUMN Quota Double"

End If

If Not ExistMDBField(CN, "PrtAssegniCircolariUNEP", "Quota") Then
   CN.Execute "ALTER TABLE  PrtAssegniCircolariUNEP ADD COLUMN Quota Double"

End If
If Not ExistMDBField(CN, "PrtEstrattoContoUNEP", "Deduzione") Then
    CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD COLUMN Deduzione Double"
End If

If Not ExistMDBField(CN, "ANAGRAFICAAVVOCATI", "PEC") Then
    CN.Execute "ALTER TABLE  ANAGRAFICAAVVOCATI ADD COLUMN PEC TEXT(100)"
    CN.Execute "ALTER TABLE  ANAGRAFICAAVVOCATI ADD COLUMN MAIL2 TEXT(100)"
End If

If Not ExistMDBTable(CN, "DEDUZIONI_UNEP") Then
  CN.Execute "CREATE TABLE DEDUZIONI_UNEP (IdCod COUNTER, CODAVV TEXT(10) NOT NULL, DataEvasionePratica TEXT(8) NOT NULL, NOTA TEXT(255), Importo Double, SIGLA TEXT(15), NumOrdinamento LONG, ANNULLO TEXT(1), Locked TEXT(50), " & _
              "CONSTRAINT pk_DEDUZIONI_UNEP_Unep PRIMARY KEY (CODAVV,DataEvasionePratica))"
  CN.Execute "INSERT INTO LockTable(TabId) VALUES('DEDUZIONI_UNEP')"

  CN.Execute "ALTER TABLE ANAGRAFICAAVVOCATI ALTER COLUMN INDIRI TEXT(100)"
  CN.Execute "ALTER TABLE ANAGRAFICAAVVOCATI ALTER COLUMN LOCALI TEXT(100)"
  CN.Execute "ALTER TABLE ANAGRAFICAAVVOCATI ALTER COLUMN NOME TEXT(100)"
  CN.Execute "ALTER TABLE ANAGRAFICAAVVOCATI ALTER COLUMN EMAIL TEXT(100)"


    Set rs = New ADODB.Recordset
    rs.Open "SELECT CODAVV, NOTE1&NOTE2&NOTE3 as Nota,EMAIL,PEC,MAIL2 FROM AnagraficaAvvocati ", CN
   Dim newMail As String
   Dim mail2 As String
   Dim pec As String
   Dim specMail As String
   Dim corretta As Boolean
    While Not rs.EOF
      codAvv = rs!codAvv
      Nota = rs!Nota
      mail = rs!EMAIL
      mail2 = IIf(IsNull(rs!mail2), "", rs!mail2)
      pec = IIf(IsNull(rs!pec), "", rs!pec)
      specMail = ""
      newMail = ""
      pos = InStr(1, Nota, "@")
      If pos > 0 Then
        'C'è una mail
        leftPart = Left(Nota, pos - 1)
        rightPart = Mid(Nota, pos)
        pos = InStrRev(leftPart, "-")
        leftPart = Mid(leftPart, pos + 1)
        pos = InStr(1, rightPart, "-")
        If pos > 0 Then rightPart = Left(rightPart, pos - 1)
        newMail = leftPart & rightPart
        pos = InStr(1, newMail, "PEC:")
        If pos > 0 Then
          newMail = Mid(newMail, pos + 4)
          specMail = newMail
        End If

        pos = InStrRev(newMail, ".")
        rightPart = Mid(newMail, pos + 1)
        corretta = Len(newMail) < 101 And InStr(1, newMail, "'") = 0 And InStr(1, newMail, " ") = 0 And InStr(1, newMail, "/") = 0 And _
                   InStr(1, newMail, "\") = 0 And InStr(1, newMail, ":") = 0 And _
                   (rightPart = "it" Or rightPart = ".com" Or rightPart = ".net" Or rightPart = ".eu")
        If corretta Then


              If Trim(mail) = "" Then


                         CN.Execute "UPDATE AnagraficaAvvocati SET EMAIL = '" & newMail & "' WHERE CODAVV='" & codAvv & "'"
                         If Trim(pec) = "" And specMail <> "" Then
                           CN.Execute "UPDATE AnagraficaAvvocati SET PEC = '" & specMail & "' WHERE CODAVV='" & codAvv & "'"
                         End If


                      Else

                         If mail <> newMail And Trim(mail2) = "" Then
                           CN.Execute "UPDATE AnagraficaAvvocati SET MAIL2 = '" & newMail & "' WHERE CODAVV='" & codAvv & "'"
                         End If
                         If Trim(pec) = "" And specMail <> "" Then
                           CN.Execute "UPDATE AnagraficaAvvocati SET PEC = '" & specMail & "' WHERE CODAVV='" & codAvv & "'"
                         End If

                      End If
              End If

      Else

        Debug.Print "!!!!ERRORE: " & newMail
      End If


      rs.MoveNext

    Wend
End If


If Not ExistMDBTable(CN, "PrtDeduzioniUNEP") Then
  CN.Execute "CREATE TABLE PrtDeduzioniUNEP (CodAvv TEXT(10),Data TEXT(10),Nota TEXT(255), Importo DOUBLE) "
End If

If Not ExistMDBField(CN, "Parametri", "CECodTestataIncasso") Then
     CN.Execute "ALTER TABLE Parametri ADD CECodTestataIncasso Text(12)"

     CN.Execute "UPDATE Parametri SET CECodTestata='100101003', CECodTestataIncasso='100101050 '"
End If

If Not ExistMDBField(CN, "TempSaldiUNEP", "Deduzione") Then

     CN.Execute "ALTER TABLE  TempSaldiUNEP ADD COLUMN Deduzione Double"
     CN.Execute "ALTER TABLE  PrtAssegniCircolariUNEP ADD COLUMN Deduzione Double"

     CN.Execute "UPDATE TempSaldiUNEP SET Deduzione=0"
     CN.Execute "UPDATE PrtAssegniCircolariUNEP SET Deduzione=0"
End If



CN.Execute "UPDATE STORICOFATTURE SET CodIVA='022' WHERE LEFT(DataFattura,4)='2016' AND CODIVA =''"
CN.Execute "UPDATE Parametri SET CodIVA='022' WHERE CODIVA =''"

If Not ExistMDBField(CN, "Parametri", "Sede") Then
     CN.Execute "ALTER TABLE Parametri ADD Sede Text(255)"

     CN.Execute "UPDATE Parametri SET Sede='Milano'"
End If

If Not ExistMDBField(CN, "PrtSospesiUNEP", "Deduzione") Then
   CN.Execute "ALTER TABLE  PrtSospesiUNEP ADD COLUMN Deduzione Double"

End If


End Sub
Public Sub UpdateDatabase(CN As ADODB.Connection)
Dim rsAv As ADODB.Recordset
Dim Ord As Long
Dim sql As String
Dim Descr As String
Dim Code As String
 Dim codAvv As String, Nota As String, mail As String, pos As Integer, leftPart As String, rightPart As String
 
 Dim rs As ADODB.Recordset
Dim nv As Integer

If Not ExistMDBTable(CN, "Versione") Then
   CN.Execute "CREATE TABLE Versione (NumeroVersione Integer) "
   CN.Execute "INSERT INTO Versione (NumeroVersione) VALUES(1)"
   CN.Execute "ALTER TABLE  PrtEstrattoConto ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  PrtEstrattoConto ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtEstrattoConto ALTER COLUMN  LOCALI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSaldi ALTER COLUMN  Commento TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSaldi ALTER COLUMN  Nome TEXT(100)"
End If

Set rs = New ADODB.Recordset
Set rs = GetADORecordset("Versione", "NumeroVersione", "1=1", CN)
If Not rs.EOF Then
  nv = rs(0)
  If nv < 2 Then
    GoTo UPD_2
   ElseIf nv < 3 Then
    GoTo UPD_3
   ElseIf nv < 4 Then
    GoTo UPD_4
   ElseIf nv < 5 Then
    GoTo UPD_5
  End If
End If
Exit Sub
UPD_2:
 Update2 CN
UPD_3:
 Update3 CN
UPD_4:
 Update4 CN
UPD_5:
 Update5 CN
End Sub
Private Sub Update5(CN As ADODB.Connection)
   CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtEstrattoContoUNEP ALTER COLUMN  LOCALI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSaldiUNEP ALTER COLUMN  Commento TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSaldiUNEP ALTER COLUMN  Nome TEXT(100)"
 
 CN.Execute "UPDATE Versione SET NumeroVersione=5"
End Sub

Private Sub Update4(CN As ADODB.Connection)
If Not ExistMDBTable(CN, "FattureTemp") Then
   CN.Execute "SELECT * INTO  FattureTemp FROM StoricoFatture "

End If

If Not ExistMDBTable(CN, "FattureTempUNEP") Then
   CN.Execute "SELECT * INTO  FattureTempUNEP FROM StoricoFattureUNEP   "

End If
 CN.Execute "UPDATE Versione SET NumeroVersione=4"
End Sub
Private Sub Update3(CN As ADODB.Connection)
 
   CN.Execute "ALTER TABLE  PrtSospesi ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSospesi ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSospesi ALTER COLUMN  LOCALI TEXT(100)"
   
   CN.Execute "ALTER TABLE  PrtSospesiUNEP ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSospesiUNEP ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  PrtSospesiUNEP ALTER COLUMN  LOCALI TEXT(100)"
 
 CN.Execute "UPDATE Versione SET NumeroVersione=3"
End Sub

Private Sub Update2(CN As ADODB.Connection)
 
   CN.Execute "ALTER TABLE  TempSaldi ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  TempSaldi ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  TempSaldi ALTER COLUMN  LOCALI TEXT(100)"
   
   CN.Execute "ALTER TABLE  TempSaldiUNEP ALTER COLUMN  Nome TEXT(100)"
   CN.Execute "ALTER TABLE  TempSaldiUNEP ALTER COLUMN  INDIRI TEXT(100)"
   CN.Execute "ALTER TABLE  TempSaldiUNEP ALTER COLUMN  LOCALI TEXT(100)"

 CN.Execute "UPDATE Versione SET NumeroVersione=2"
End Sub


