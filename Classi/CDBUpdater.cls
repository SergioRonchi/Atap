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

Public Sub UpdateDatabase(cn As ADODB.Connection)
Dim rsAv As ADODB.Recordset
Dim Ord As Long, codAvv As String
Dim SQL As String
Dim Descr As String
Dim code As String
If Not ExistMDBTable(cn, "LockTable") Then
     cn.Execute "CREATE TABLE LockTable (TabID TEXT(50) PRIMARY KEY, Locked TEXT(50) DEFAULT 'NONE')"
     cn.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('ADEMPI','NONE')"
     cn.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('AnagraficaAvvocati','NONE')"
     cn.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('DecretiIngiuntivi','NONE')"
     cn.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('Notifiche','NONE')"
     cn.Execute "INSERT INTO LockTable (TabID, Locked) VALUES ('SFRATTI','NONE')"
End If
If Not ExistMDBTable(cn, "NumeriFattura") Then
     cn.Execute "CREATE TABLE NumeriFattura (Numero LONG)"
     cn.Execute "INSERT INTO NumeriFattura (Numero) VALUES (1)"
End If
If Not ExistMDBField(cn, "Anticipi", "Locked") Then
     cn.Execute "ALTER TABLE Anticipi ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE Anticipi SET Locked='NONE'"
 End If
 
If Not ExistMDBField(cn, "ADEMPI", "Locked") Then
     cn.Execute "ALTER TABLE ADEMPI ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE ADEMPI SET Locked='NONE'"
     cn.Execute "ALTER TABLE AnagraficaAvvocati ADD COLUMN Locked TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE  AnagraficaAvvocati SET Locked='NONE'"
     cn.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE  DecretiIngiuntivi SET Locked='NONE'"
     cn.Execute "ALTER TABLE Notifiche ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE  Notifiche SET Locked='NONE'"
     cn.Execute "ALTER TABLE SFRATTI ADD COLUMN Locked  TEXT(50) DEFAULT 'NONE'"
     cn.Execute "UPDATE  SFRATTI SET Locked='NONE'"

End If
'Rivedi ordinamento avvocati
If Not ExistMDBField(cn, "Parametri", "XXX") Then
   'Cerca dupplicati
   SQL = "SELECT Anticipi.Descrizione, Anticipi.CodiceAnticipi, Anticipi_1.CodiceAnticipi " & _
       "FROM Anticipi, Anticipi AS Anticipi_1 " & _
       "WHERE (((Anticipi.Descrizione)=[Anticipi_1].[Descrizione]) AND ((Anticipi.CodiceAnticipi)>[Anticipi_1].[CodiceAnticipi]));"
   Set rsAv = newAdoRs
   rsAv.Open SQL, cn
   While Not rsAv.EOF
     code = rsAv(1)
     Descr = rsAv(0) & " " & code
     cn.Execute "UPDATE ANTICIPI SET DESCRIZIONE='" & Replace(Descr, "'", "''") & "' WHERE CodiceAnticipi='" & Replace(code, "'", "''") & "'; "
     rsAv.MoveNext
   Wend
   rsAv.Close
  cn.Execute "CREATE UNIQUE INDEX uniDescr ON ANTICIPI (Descrizione)"
  cn.Execute "ALTER TABLE Parametri ADD COLUMN XXX logical"
  Set rsAv = GetADORecordset("AnagraficaAvvocati", "CodAvv", "1=1 ORDER BY NumOrdinamento", cn)
  OpenProgress ("Attendere... Aggiornamento in corso!")
  While Not rsAv.EOF
   UpdateProgress rsAv.AbsolutePosition * 100 / rsAv.RecordCount
   Ord = Ord + 50
   codAvv = rsAv(0)
   cn.Execute "UPDATE AnagraficaAvvocati set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE ADEMPI set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE DecretiIngiuntivi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE Notifiche set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtAssegniCircolari set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtEstrattoConto set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtGiornalieraAdempimenti set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtGiornalieraDecretiIngiuntivi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtGiornalieraNotifiche set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtGiornalieraSfrattiPig set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE PrtSaldi set NumOrdinamento=" & Ord & " WHERE Codice='" & codAvv & "'"
   cn.Execute "UPDATE PrtSospesi set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE Saldi set NumOrdinamento=" & Ord & " WHERE Codice='" & codAvv & "'"
   cn.Execute "UPDATE Sfratti set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   cn.Execute "UPDATE StoricoFatture set NumOrdinamento=" & Ord & " WHERE CodAvv='" & codAvv & "'"
   rsAv.MoveNext
  Wend
  CloseProgress
  
  
End If
cn.Execute "UPDATE ANAGRAFICAAVVOCATI SET CASSETTAROTTA='N' WHERE CASSETTAROTTA Is Null"
cn.Execute "UPDATE ANAGRAFICAAVVOCATI SET STAT='V' WHERE STAT Is Null"
If Not ExistMDBField(cn, "prtEstrattoConto", "CodAutorita") Then
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN CodAutorita VARCHAR(2)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Esenzione VARCHAR(1)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN FormulaEsec VARCHAR(1)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN NumeroIngiunzione VARCHAR(6)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN NumeroRuolo VARCHAR(7)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Localita1 VARCHAR(35)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Nota VARCHAR(14)"
  cn.Execute "ALTER TABLE prtEstrattoConto ADD COLUMN Descrizione VARCHAR(50)"
  cn.Execute "ALTER TABLE prtSospesi ADD COLUMN Nota VARCHAR(14)"
End If
If Not ExistMDBField(cn, "DecretiIngiuntivi", "IDCod") Then
  cn.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN IDCod COUNTER"
  cn.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN Ricorrente Text(35)"
  cn.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN Debitore Text(35)"
  cn.Execute "ALTER TABLE DecretiIngiuntivi ALTER COLUMN DesrSpese Text(35)"
End If
If Not ExistMDBField(cn, "Parametri", "IVA") Then
  cn.Execute "ALTER TABLE Parametri ADD COLUMN IVA DOUBLE"
  cn.Execute "UPDATE Parametri SET IVA=20.0"
End If
If Not ExistMDBField(cn, "AnagraficaAvvocati", "ID") Then
  cn.Execute "ALTER TABLE AnagraficaAvvocati ADD COLUMN ID COUNTER"
  cn.Execute "ALTER TABLE AnagraficaAvvocati ALTER COLUMN NOME Text(50)"
End If
If Not ExistMDBField(cn, "ADEMPI", "SIGLA") Then
  cn.Execute "ALTER TABLE ADEMPI ADD COLUMN SIGLA VARCHAR(15)"
  cn.Execute "ALTER TABLE Sfratti ADD COLUMN SIGLA VARCHAR(15)"
  cn.Execute "ALTER TABLE Notifiche ADD COLUMN SIGLA VARCHAR(15)"
  cn.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN SIGLA VARCHAR(15)"
  cn.Execute "ALTER TABLE ADEMPI ADD COLUMN SIGLACH VARCHAR(15)"
  cn.Execute "ALTER TABLE Sfratti ADD COLUMN SIGLACH VARCHAR(15)"
  cn.Execute "ALTER TABLE Notifiche ADD COLUMN SIGLACH VARCHAR(15)"
  cn.Execute "ALTER TABLE DecretiIngiuntivi ADD COLUMN SIGLACH VARCHAR(15)"
  
  cn.Execute "UPDATE DecretiIngiuntivi SET SIGLACH='OLD'"
  cn.Execute "UPDATE ADEMPI SET SIGLACH='OLD'"
  cn.Execute "UPDATE Sfratti SET SIGLACH='OLD'"
  cn.Execute "UPDATE Notifiche SET SIGLACH='OLD'"

  cn.Execute "UPDATE DecretiIngiuntivi SET SIGLA='OLD'"
  cn.Execute "UPDATE ADEMPI SET SIGLA='OLD'"
  cn.Execute "UPDATE Sfratti SET SIGLA='OLD'"
  cn.Execute "UPDATE Notifiche SET SIGLA='OLD'"

End If
If Not ExistMDBField(cn, "Notifiche", "IDCod") Then
  cn.Execute "ALTER TABLE Notifiche ADD COLUMN IDCod COUNTER"
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN NaturaAtto Text(35)"
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte1 Text(35)"
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte2 Text(35)"
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN DesrSpese Text(35)"
  cn.Execute "ALTER TABLE Sfratti ALTER COLUMN Localita1 Text(40)"
  
End If
If Not ExistMDBField(cn, "PrtSospesi", "DATA_EVASIONE") Then
  cn.Execute "ALTER TABLE PrtSospesi ADD DATA_EVASIONE Date"
  
  cn.Execute "ALTER TABLE PrtSospesi ADD SALDO_PRECEDENTE Double"
  cn.Execute "ALTER TABLE PrtSospesi ADD NUM_EST_CONTO Long"
  cn.Execute "ALTER TABLE PrtSospesi ADD DATA_EST_CONTO DAte"
  cn.Execute "ALTER TABLE PrtSospesi ADD PROVVISORIO Text(1)"
  cn.Execute "ALTER TABLE PrtSospesi ADD DATA_INIZIO Text(10)"
  cn.Execute "ALTER TABLE PrtSospesi ADD DATA_FINE Text(10)"
    
  cn.Execute "ALTER TABLE PrtSospesi ADD QtaMarche Long"
  cn.Execute "ALTER TABLE PrtSospesi ADD QtaFotocopie Long"
  cn.Execute "ALTER TABLE PrtSospesi ADD QtaCopie Text(5)"
  cn.Execute "ALTER TABLE PrtSospesi ADD QtaDirittiCancelleria Long"
     
  End If

If Not ExistMDBField(cn, "Sfratti", "IDCod") Then
  cn.Execute "ALTER TABLE Sfratti ADD COLUMN IDCod COUNTER"
  cn.Execute "ALTER TABLE Sfratti ALTER COLUMN Localita1 Text(40)"
  cn.Execute "ALTER TABLE Sfratti ALTER COLUMN Parte1 Text(35)"
  cn.Execute "ALTER TABLE Sfratti ALTER COLUMN Parte2 Text(35)"
  cn.Execute "ALTER TABLE Sfratti ALTER COLUMN DesrSpese Text(35)"
End If
If Not ExistMDBField(cn, "ADEMPI", "IDCod") Then
  cn.Execute "ALTER TABLE ADEMPI ADD COLUMN IDCod COUNTER"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese1 Text(35)"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese2 Text(35)"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese3 Text(35)"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese4 Text(35)"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese5 Text(35)"
  cn.Execute "ALTER TABLE ADEMPI ALTER COLUMN DesrSpese6 Text(35)"
  
  cn.Execute "ALTER TABLE PrtAssegniCircolari ALTER COLUMN LOCALI Text(40)"
  cn.Execute "ALTER TABLE PrtAssegniCircolari ALTER COLUMN Nome Text(50)"
  
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN LOCALI Text(40)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Nome Text(50)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Parte1 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN Parte2 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE1 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE2 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE3 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE4 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE5 Text(35)"
  cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DESCR_SPESE6 Text(35)"
  
  cn.Execute "ALTER TABLE PrtFattProv ALTER COLUMN LOCALI Text(40)"
  cn.Execute "ALTER TABLE PrtFattProv ALTER COLUMN Nome Text(50)"
   
  cn.Execute "ALTER TABLE PrtGiornalieraDecretiIngiuntivi ALTER COLUMN Ricorrente Text(35)"
  cn.Execute "ALTER TABLE PrtGiornalieraDecretiIngiuntivi ALTER COLUMN Debitore Text(35)"
  
  cn.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN parte1 Text(35)"
  cn.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN parte2 Text(35)"
  cn.Execute "ALTER TABLE PrtGiornalieraNotifiche ALTER COLUMN Localita1 Text(40)"
  
  cn.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN parte1 Text(35)"
  cn.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN parte2 Text(35)"
  cn.Execute "ALTER TABLE PrtGiornalieraSfrattiPig ALTER COLUMN Localita1 Text(40)"
  
  cn.Execute "ALTER TABLE PrtSaldi ALTER COLUMN Nome Text(50)"
  
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN Nome Text(50)"
  
  
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN LOCALI Text(40)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN PARTE1 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN PARTE2 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE1 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE2 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE3 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE4 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE5 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DESCR_SPESE6 Text(35)"
  cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN Localita1 Text(35)"
  
'Non rimuovo per il momento le tabelle
'  cn.Execute "DROP TABLE Filter"
'  cn.Execute "DROP TABLE LockAdempimenti"
'  cn.Execute "DROP TABLE LockDecreti"
'  cn.Execute "DROP TABLE LockNotifiche"
'  cn.Execute "DROP TABLE LockPignoramenti"

End If

If Not ExistMDBTable(cn, "LockPrt") Then
  cn.Execute "CREATE TABLE LockPrt (" & _
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
  cn.Execute "INSERT INTO LockPrt(PrtAssegniCircolari," & _
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
If Not ExistMDBField(cn, "PrtEStrattoConto", "DataRegistrazione") Then
   cn.Execute "ALTER TABLE PrtEstrattoConto ADD COLUMN DataRegistrazione Text(10) NOT NULL"
   cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DATA_EVASIONE Text(10) NOT NULL"
   cn.Execute "ALTER TABLE PrtSospesi ALTER COLUMN DATARegistrazione Text(10) NOT NULL"
   cn.Execute "ALTER TABLE PrtEstrattoConto ALTER COLUMN DATA_EVASIONE Text(10) NOT NULL"
End If
If Not ExistMDBTable(cn, "TempSaldi") Then
     cn.Execute "SELECT * INTO TempSaldi FROM PrtAssegniCircolari "
     cn.Execute "DELETE * FROM TempSaldi"
End If

'Elimina duplicati nei saldi
Dim eliminati As String
Dim key As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT Saldi.Codice, Saldi.Chiusura, Saldi.SaldoAdemp, Saldi.SaldoAdempEuro, Saldi.SaldoSfpg, Saldi.SaldoSfpgEuro, Saldi.SaldoNotif, Saldi.SaldoNotifEuro, Saldi.SaldoDecrIng, Saldi.SaldoDecrIngEuro, Saldi.Stato, Saldi.SaldoTotale, Saldi.SaldoTotaleEuro, Saldi.PROG_Saldi, Saldi.Commento, Saldi.NumOrdinamento " & _
        " FROM Saldi INNER JOIN Saldi AS Saldi_1 ON Saldi.Codice = Saldi_1.Codice" & _
        " WHERE (((Saldi.CHIUSURA)<>[Saldi_1].[cHIUSURA])) ORDER BY  Saldi.CHIUSURA DESC;", cn
While Not rs.EOF
  key = "$" & rs!codice & "$"
  If InStr(1, eliminati, key) = 0 Then
    cn.Execute "DELETE * FROM Saldi WHERE Codice='" & rs!codice & "' AND Chiusura<>'" & rs!Chiusura & "'"
    eliminati = eliminati + key
  End If
  
  rs.MoveNext
  
Wend

'-----------------------------------------------------------------------------------------------------------------------------------------
If Not ExistMDBField(cn, "Sfratti", "Crono") Then
  cn.Execute "ALTER TABLE Sfratti ADD COLUMN IsUNEP Logical"
  cn.Execute "UPDATE Sfratti SET  IsUNEP=false"
  
  cn.Execute "ALTER TABLE Sfratti ADD COLUMN Crono TEXT(50)"
  
End If

If Not ExistMDBField(cn, "Notifiche", "Crono") Then
  cn.Execute "ALTER TABLE Notifiche ADD COLUMN IsUNEP Logical"
  cn.Execute "UPDATE Notifiche SET  IsUNEP=false"
  
  cn.Execute "ALTER TABLE Notifiche ADD COLUMN Crono TEXT(50)"
  
End If

If Not ExistMDBField(cn, "Parametri", "QuotaSoci") Then
  cn.Execute "ALTER TABLE Parametri ADD COLUMN QuotaSoci Double"
  cn.Execute "UPDATE Parametri SET  QuotaSoci=55"
    
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte1 VARCHAR(35) NULL"
  cn.Execute "ALTER TABLE Notifiche ALTER COLUMN Parte2 VARCHAR(35) NULL"
  
  cn.Execute "CREATE TABLE PrtGiornalieraNotificheUNEP (Crono TEXT(50), DescrizioneTribunale TEXT(30), CODAVV TEXT(10), NumeroAtto LONG, DataRegistrazione TEXT(10), " & _
             "parte1 TEXT(35), parte2 TEXT(35), [Note] TEXT(14), Localita1 TEXT(40), Descrizione TEXT(50), ImpSaldo DOUBLE, Valuta TEXT(4), NumOrdinamento LONG, " & _
             "DataPresentazione TEXT(10), DataRestituzione TEXT(10), DataNotifica TEXT(10), DATA_INIZIO TEXT(10), DATA_FINE TEXT(10))"
             
  cn.Execute "CREATE TABLE PrtGiornalieraSfrattiPigUNEP (Crono TEXT(50), DescrizioneTribunale TEXT(30), CODAVV TEXT(10), NumeroAtto LONG, DataRegistrazione TEXT(10), DataPresentazione TEXT(10)," & _
             "parte1 TEXT(35), parte2 TEXT(35), Localita1 TEXT(40), Descrizione TEXT(50), ImpSaldo DOUBLE, Valuta TEXT(4), NumOrdinamento LONG, " & _
              "DataRestituzione TEXT(10),  DATA_INIZIO TEXT(10), DATA_FINE TEXT(10))"
             
  cn.Execute "ALTER TABLE LockPrt ADD PrtGiornalieraNotificheUNEP Logical, PrtGiornalieraSfrattiPigUNEP Logical"
  cn.Execute "UPDATE LockPrt SET PrtGiornalieraNotificheUNEP=false, PrtGiornalieraSfrattiPigUNEP=false"
End If

Set rs = New ADODB.Recordset
rs.Open "SELECT CodiceTribunale FROM TribunaliAppartenenza WHERE CodiceTribunale='UNEP'", cn

If rs.EOF Then
  cn.Execute "INSERT INTO TribunaliAppartenenza (CodiceTribunale, DescrizioneTribunale) VALUES('UNEP', 'MILANO UNEP')"
End If

If Not ExistMDBTable(cn, "PrtEstrattoContoUNEP") Then
  cn.Execute "SELECT * INTO PrtEstrattoContoUNEP FROM PrtEstrattoConto"
  
End If

If Not ExistMDBTable(cn, "TempSaldiUNEP") Then
  cn.Execute "SELECT * INTO TempSaldiUNEP FROM TempSaldi"
  
End If

If Not ExistMDBTable(cn, "PrtAssegniCircolariUNEP") Then
  cn.Execute "SELECT * INTO PrtAssegniCircolariUNEP FROM PrtAssegniCircolari"
  
End If

If Not ExistMDBTable(cn, "StoricoFattureUNEP") Then
  cn.Execute "SELECT * INTO StoricoFattureUNEP FROM StoricoFatture"
  cn.Execute "DELETE * FROM StoricoFattureUNEP"
End If


If Not ExistMDBTable(cn, "SaldiUNEP") Then
 
  cn.Execute "SELECT * INTO SaldiUNEP FROM Saldi"
  cn.Execute "UPDATE SaldiUNEP SET SaldoAdemp=0, SaldoAdempEuro=0,  SaldoSfpg=0, SaldoSfpgEuro=0, SaldoNotif=0, SaldoNotifEuro=0, SaldoDecrIng=0, SaldoDecrIngEuro=0, SaldoTotale=0, SaldoTotaleEuro=0, Commento=''"
End If

If Not ExistMDBTable(cn, "PrtSaldiUNEP") Then
  cn.Execute "SELECT * INTO PrtSaldiUNEP FROM PrtSaldi"
  cn.Execute "DELETE * FROM PrtSaldiUNEP"
End If

If Not ExistMDBTable(cn, "PrtSospesiUNEP") Then
  cn.Execute "SELECT * INTO PrtSospesiUNEP FROM PrtSospesi"
  cn.Execute "DELETE * FROM PrtSospesiUNEP"
End If


If Not ExistMDBField(cn, "LockPrt", "PrtAssegniCircolariUNEP") Then
  cn.Execute "ALTER TABLE  LockPrt ADD PrtAssegniCircolariUNEP Logical, PrtEstrattoContoUNEP Logical"
  
  cn.Execute "UPDATE LockPrt SET PrtAssegniCircolariUNEP=false, PrtEstrattoContoUNEP=false "
End If

If Not ExistMDBField(cn, "PrtEstrattoContoUNEP", "DATA_PRESENTAZIONE") Then

  cn.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD Crono Text(50)"
  
  
  cn.Execute "ALTER TABLE  PrtEstrattoConto ADD Crono Text(50)"
  
  
  cn.Execute "ALTER TABLE  prtSospesi ADD Crono Text(50)"

  cn.Execute "ALTER TABLE  PrtEstrattoContoUNEP ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"
  
  
  cn.Execute "ALTER TABLE  PrtEstrattoConto ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"
  
  
  cn.Execute "ALTER TABLE  prtSospesi ADD DATA_PRESENTAZIONE Text(10), DATA_RESTITUZIONE Text(10)"
 
End If

If Not ExistMDBTable(cn, "PrtData") Then
   cn.Execute "CREATE TABLE PrtData (Bimestre Integer, BimestreAnno Integer)"
  
End If


If Not ExistMDBField(cn, "LockPrt", "PrtSaldiUNEP") Then
  cn.Execute "ALTER TABLE  LockPrt ADD PrtSaldiUNEP Logical, PrtSospesiUNEP Logical"
  
  cn.Execute "UPDATE LockPrt SET PrtSaldiUNEP=false, PrtSospesiUNEP=false "
End If

If Not ExistMDBField(cn, "Date_EstrattiConto", "DATA_ULTIMO_ESTCONTO_UNEP") Then
  cn.Execute "ALTER TABLE  Date_EstrattiConto ADD DATA_ULTIMO_ESTCONTO_UNEP TEXT(10), DATA_FATTURA_UNEP TEXT(10)"

End If

If Not ExistMDBField(cn, "StoricoFattureUNEP", "Bimestre") Then
  cn.Execute "ALTER TABLE  StoricoFattureUNEP ADD Bimestre TEXT(50), QUOTA DOUBLE"

End If
'-----------------------------------------------------------------------------------------------

If Not ExistMDBTable(cn, "SFRATTI_UNEP") Then
  cn.Execute "SELECT * INTO SFRATTI_UNEP FROM SFRATTI"
  cn.Execute "DELETE * FROM SFRATTI_UNEP WHERE IsUNEP=False"
  cn.Execute "DELETE * FROM SFRATTI WHERE IsUNEP=True"
  
  cn.Execute "ALTER TABLE  SFRATTI DROP COLUMN IsUNEP"
  cn.Execute "ALTER TABLE  SFRATTI_UNEP DROP COLUMN IsUNEP"
  cn.Execute "INSERT INTO LockTable(TabId) VALUES('SFRATTI_UNEP')"
End If

If Not ExistMDBTable(cn, "NOTIFICHE_UNEP") Then
  cn.Execute "SELECT * INTO NOTIFICHE_UNEP FROM NOTIFICHE"
  cn.Execute "DELETE * FROM NOTIFICHE_UNEP WHERE IsUNEP=False"
  cn.Execute "DELETE * FROM NOTIFICHE WHERE IsUNEP=True"
  
  cn.Execute "ALTER TABLE  NOTIFICHE DROP COLUMN IsUNEP"
  cn.Execute "ALTER TABLE  NOTIFICHE_UNEP DROP COLUMN IsUNEP"
  cn.Execute "INSERT INTO LockTable(TabId) VALUES('NOTIFICHE_UNEP')"
  
End If

If Not ExistMDBField(cn, "PrtData", "Tipo") Then
  cn.Execute "ALTER TABLE  PrtData ADD Tipo Integer"

End If

If Not ExistMDBField(cn, "Parametri", "LimiteSaldo") Then
  cn.Execute "ALTER TABLE  Parametri ADD LimiteSaldo DOUBLE"
  cn.Execute "UPDATE Parametri SET LimiteSaldo=5.16456"

End If
If Not ExistMDBField(cn, "Parametri", "Banca") Then
  cn.Execute "ALTER TABLE  Parametri ADD Banca TEXT(200), IBAN TEXT(50)"
  cn.Execute "UPDATE Parametri SET Banca='BNL AG. 4373 ', IBAN='IT 45 J 01005 01773 000000220006'"

End If

If Not ExistMDBField(cn, "Parametri", "Telefono") Then
  cn.Execute "ALTER TABLE  Parametri ADD COLUMN Telefono TEXT(20)"
  cn.Execute "ALTER TABLE  Sfratti ALTER COLUMN CodicePignoramenti TEXT(5)"
  cn.Execute "ALTER TABLE  Sfratti_UNEP ALTER COLUMN CodicePignoramenti TEXT(5)"

  cn.Execute "ALTER TABLE Sfratti_Unep ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,DataRegistrazione,NumeroAtto,CodTribunaleApp)"
  cn.Execute "ALTER TABLE StoricoFattureUNEP ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,NumeroFattura,DataFattura)"
  cn.Execute "ALTER TABLE Notifiche_UNEP ADD CONSTRAINT pk_Sfratti_Unep PRIMARY KEY (CODAVV,DataRegistrazione,NumeroAtto,CodTribunaleApp)"
  
End If



End Sub



