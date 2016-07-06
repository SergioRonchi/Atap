Attribute VB_Name = "mdlEstrattiConto"
Option Explicit
Public Function isToTransfer(Tabella As String, schema As String) As Boolean
 isToTransfer = True
 Tabella = UCase(Tabella)
 
 If Tabella = "ADEMPI" Then isToTransfer = (InStr(1, schema, "A") <> 0)
 If Tabella = "NOTIFICHE" Then isToTransfer = (InStr(1, schema, "N") <> 0)
 If Tabella = "SFRATTI" Then isToTransfer = (InStr(1, schema, "S") <> 0)
 If Tabella = "DECRETIINGIUNTIVI" Then isToTransfer = (InStr(1, schema, "D") <> 0)

 If Tabella = "NOTIFICHE_UNEP" Then isToTransfer = (InStr(1, schema, "N") <> 0)
 If Tabella = "SFRATTI_UNEP" Then isToTransfer = (InStr(1, schema, "S") <> 0)
End Function
Public Function GetFreeFile(NomeFile As String) As String
 Dim i As Integer
 Dim s As String
 i = 1
 s = NomeFile
 While Dir(s) <> ""
   s = NomeFile & "_" & i
   i = i + 1
 Wend
 GetFreeFile = s
End Function
Public Function Trasferisci(ByRef NomeFile As String, Da As String, A As String, isUnep As Boolean, Optional codAvv As String, Optional schema As String) As Boolean
  Dim rsTable As ADODB.Recordset
  Dim SQL As String, Tabella As String
  Dim sqlDEL As String
  Dim anno As String
  Dim Data As Boolean
  Dim isUnepTable As Boolean
  Dim sWhere As String
  Dim isCodAvv As Boolean
  
  On Error GoTo errtrasf
  Screen.MousePointer = vbHourglass
  NomeFile = GetFreeFile(NomeFile)
  
  Dim portion2 As String
  Dim portion1 As String
  
  Dim p As Integer
  p = InStr(1, NomeFile, "EC_")
  If p < 1 Then
     p = InStr(1, NomeFile, "LIQ")
  End If
  If p < 1 Then
     p = InStr(1, NomeFile, "ECUNEP_")
  End If
  If p < 1 Then
     p = InStr(1, NomeFile, "LIQUNEP")
  End If
  
  
  portion1 = Left(NomeFile, p)
  portion2 = Mid(NomeFile, p + 1)
  
  portion2 = Replace(portion2, "/", "_")
  NomeFile = portion1 & Replace(portion2, "\", "_")
  
  If Dir(NomeFile) = "" Then
      FileCopy app.Path & "\sto.0", NomeFile
      
      'Creazione dei dati nel file STORICO
      g_Settings.DBConnection.BeginTrans
      Set rsTable = g_Settings.DBConnection.OpenSchema(adSchemaColumns)
     
      While Not rsTable.EOF
        
          Tabella = rsTable!TABLE_NAME
          isUnepTable = UCase(Tabella) = "SFRATTI_UNEP" Or UCase(Tabella) = "NOTIFICHE_UNEP"
          If UCase(Left(Tabella, 3)) <> "QRY" Then
                    Data = False
                    isCodAvv = False
                    
                    If Left(Tabella, 4) <> "MSys" And UCase(Left(Tabella, 4)) <> "TEMP" And UCase(Tabella) <> "TAB_NOTE" And UCase(Left(Tabella, 4)) <> "~TMP" And UCase(Left(Tabella, 3)) <> "qry" Then
                            Do
                              If UCase(rsTable!COLUMN_NAME) = "DATAEVASIONEPRATICA" Then Data = True
                              If UCase(rsTable!COLUMN_NAME) = "CODAVV" Then isCodAvv = True
                              
                              rsTable.MoveNext
                              If rsTable.EOF Then Exit Do
                            Loop While rsTable!TABLE_NAME = Tabella
                            
                            If Data Then
                                SQL = "SELECT *  INTO [" & Tabella & "] IN '" & NomeFile & "' FROM [" & Tabella & "]" & _
                                          " WHERE DataEvasionePratica>='" & Da & "' AND DataEvasionePratica<='" & A & "' "
                                If (isUnep And isUnepTable) Or (Not isUnep And Not isUnepTable) Then

                                        sqlDEL = "DELETE * FROM [" & Tabella & "] WHERE DataEvasionePratica>='" & Da & "' AND DataEvasionePratica<='" & A & "'"
                                   Else
                                          
                                        sqlDEL = ""
                                End If
               
                               
                              Else
                                SQL = "SELECT *  INTO [" & Tabella & "] IN '" & NomeFile & "' FROM [" & Tabella & "]"
                                sqlDEL = ""
                            End If
                            If isCodAvv And codAvv <> "" Then
                               If Data Then
                                 SQL = SQL & " AND [" & Tabella & "].CodAvv='" & codAvv & "'"
                                    If (isUnep And isUnepTable) Or (Not isUnep And Not isUnepTable) Then
                                        sqlDEL = sqlDEL & " AND [" & Tabella & "].CodAvv='" & codAvv & "'"
                                    Else
                                        sqlDEL = ""
                                    End If
                                   Else
                                    SQL = SQL & " WHERE [" & Tabella & "].CodAvv='" & codAvv & "'"
                                    sqlDEL = ""
                               End If
                            End If
                            
                            If Not isToTransfer(Tabella, schema) Then
                              SQL = SQL & " AND [" & Tabella & "].CodAvv='XXXXX'"
                              If sqlDEL <> "" Then sqlDEL = sqlDEL & " AND [" & Tabella & "].CodAvv='XXXXX'"
                            End If
                            g_Settings.DBConnection.Execute SQL
                            
                            If sqlDEL <> "" Then
                               g_Settings.DBConnection.Execute sqlDEL
                            End If
                     End If
            
       
        End If
        If Not rsTable.EOF Then rsTable.MoveNext
      Wend
      g_Settings.DBConnection.CommitTrans
      rsTable.Close
      MsgBox "Trasferimento nel database" & vbCrLf & NomeFile & vbCrLf & "Eseguito Correttamente", vbOKOnly + vbInformation
      Trasferisci = True
      
    Else
     MsgBox "Il file" & vbCrLf & NomeFile & vbCrLf & "Esiste già, impossibile continuare."
     Trasferisci = False
  End If
  Screen.MousePointer = vbDefault
  Exit Function
errtrasf:
  MsgBox err.Description & vbCrLf & SQL, vbOKOnly + vbCritical
  g_Settings.DBConnection.RollbackTrans
End Function

Public Function getQryAdempimenti(isUnep As Boolean, tipo As String, Da As String, A As String, provvisorio As String, Optional Sospeso As Boolean) As String
    ' Adempimenti
    Dim qrySQL As String
    Dim sSaldi As String
    
    If isUnep Then
      sSaldi = "SaldiUNEP"
    Else
     sSaldi = "Saldi"
    End If
    qrySQL = " SELECT AnagraficaAvvocati.CODAVV, AnagraficaAvvocati.NOME, AnagraficaAvvocati.INDIRI, "
    qrySQL = qrySQL & " AnagraficaAvvocati.LOCALI, AnagraficaAvvocati.PROV,"
    qrySQL = qrySQL & "  AnagraficaAvvocati.CAP , AnagraficaAvvocati.TELEFCELL, AnagraficaAvvocati.TELEF,  "
    qrySQL = qrySQL & " progressivo,ImpDepositoE, ImpSpese1E, DesrSpese1, "
    qrySQL = qrySQL & " ImpSpese2E, DesrSpese2, ImpSpese3E,DesrSpese3, "
    qrySQL = qrySQL & " ImpSpese4E, DesrSpese4,  "
    qrySQL = qrySQL & " ImpSpese5E, DesrSpese5, ImpSpese6E, DesrSpese6, "
    qrySQL = qrySQL & " ImpCompetenzeE, "
    If tipo = "Futuro" Then
       qrySQL = qrySQL & " ImpDepositoE-ImpSpese1E-ImpSpese2E-ImpSpese3E-ImpSpese4E-ImpSpese5E-ImpSpese6E-ImpCompetenzeE*" & Str(1 + g_Settings.IVA / 100) & " AS SaldoFinale, "
      Else
       qrySQL = qrySQL & " ImpSaldoE  AS SaldoFinale, "
    End If
    
    qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
    qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
    
    If Sospeso Then
          qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
          qrySQL = qrySQL & " DataRegistrazione as D , AttivitaRichiesta, "

     Else
          qrySQL = qrySQL & " Mid(DataEvasionePratica,7,2) & '/' &  Mid(DataEvasionePratica,5,2) & '/' &  Mid(DataEvasionePratica,1,4), "
          qrySQL = qrySQL & " (MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4)) as D , AttivitaRichiesta, "

    End If
    
    qrySQL = qrySQL & " TribunaliAppartenenza.DescrizioneTribunale,  "
    qrySQL = qrySQL & " " & sSaldi & ".SaldoTotaleEuro,  " & sSaldi & ".PROG_Saldi+1 as Num, '" & Format(Date, "dd/mm/yyyy") & "' as Data,"
    qrySQL = qrySQL & " 'E' as Valuta,'" & provvisorio & "' as Provvisorio,'" & Da & "' as DATA_INIZIO, '" & A & "' as DATA_FINE,"
    qrySQL = qrySQL & "'' as Parte1,'' as Parte2, 'Adempimenti Cancelleria' as DesACT,AnagraficaAvvocati.NumOrdinamento "
    qrySQL = qrySQL & "FROM Parametri," & sSaldi & " INNER JOIN ((AnagraficaAvvocati INNER JOIN ADEMPI ON AnagraficaAvvocati.CODAVV = ADEMPI.CODAVV) "
    qrySQL = qrySQL & "INNER JOIN TribunaliAppartenenza ON ADEMPI.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) "
    qrySQL = qrySQL & "ON " & sSaldi & ".Codice = AnagraficaAvvocati.CODAVV "
    
    If Sospeso Then
          qrySQL = qrySQL & " WHERE (((ADEMPI.DataEvasionePratica)='') AND ((ADEMPI.Annullo)='V'))"
     Else
          qrySQL = qrySQL & " WHERE (((ADEMPI.DataEvasionePratica)<>'') AND ((ADEMPI.Annullo)='V'))"
    End If
    

    
    getQryAdempimenti = qrySQL
End Function


Public Function getQryDecreti(isUnep As Boolean, tipo As String, Da As String, A As String, provvisorio As String, Optional Sospeso As Boolean) As String
    ' Decreti
    Dim qrySQL As String
    Dim sSaldi As String
    
    If isUnep Then
      sSaldi = "SaldiUNEP"
    Else
     sSaldi = "Saldi"
    End If
    qrySQL = " SELECT AnagraficaAvvocati.CODAVV, AnagraficaAvvocati.NOME, AnagraficaAvvocati.INDIRI, AnagraficaAvvocati.LOCALI, "
    qrySQL = qrySQL & "AnagraficaAvvocati.PROV, AnagraficaAvvocati.CAP, AnagraficaAvvocati.TELEFCELL, AnagraficaAvvocati.TELEF, NumeroDecreto, "
    qrySQL = qrySQL & "ImpFotocopieE, 'Fotocopie', "
    qrySQL = qrySQL & "ImpFormulaE, 'Costo Formula', ImpMarcheE,'Marche', "
    qrySQL = qrySQL & "ImpSpeseE,  DesrSpese, ImpCopieE,'Diritti Cancelleria',0,' ',ImpCompetenzeE,  "
    If tipo = "Futuro" Then
        qrySQL = qrySQL & "  ImpDepositoE - (ImpFotocopieE * QtaFotocopie) - (ImpMarcheE * QtaMarche) - (ImpCopieE * QtaDirittiCancelleria) - ImpFormulaE - ImpSpeseE  - ImpCompetenzeE*" & Str(1 + g_Settings.IVA / 100) & "   AS SaldoFinale, "
      Else
       qrySQL = qrySQL & " ImpSaldoE  AS SaldoFinale, "
    End If
    
    qrySQL = qrySQL & " MID(DataDecreto,7,2) & '/' & MID(DataDecreto,5,2)& '/' & MID(DataDecreto,1,4), "
    qrySQL = qrySQL & " MID(DataDecreto,7,2) & '/' & MID(DataDecreto,5,2)& '/' & MID(DataDecreto,1,4), "
    
     If Sospeso Then
          qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
          qrySQL = qrySQL & " DataRegistrazione as D , ' ', "


     Else
          qrySQL = qrySQL & " Mid(DataEvasionePratica,7,2) & '/' &  Mid(DataEvasionePratica,5,2) & '/' &  Mid(DataEvasionePratica,1,4), "
          qrySQL = qrySQL & " (MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4)) as D , ' ', "

    End If
   qrySQL = qrySQL & "TribunaliAppartenenza.DescrizioneTribunale," & sSaldi & ".SaldoTotaleEuro," & sSaldi & ".PROG_Saldi+1,'" + Format(Now, "dd/mm/yyyy") + "',"
    qrySQL = qrySQL & "'E','" & provvisorio & "','" & Da & "','" & A & "',  Ricorrente, Debitore,'Decreti Ingiuntivi',  "
    qrySQL = qrySQL & "QtaCopie,QtaFotocopie,QtaMarche,QtaDirittiCancelleria,ImpDepositoE,AnagraficaAvvocati.NumOrdinamento,CodAutorita,Esenzione,FormulaEsec,NumeroIngiunzione,NumeroRuolo  "
    qrySQL = qrySQL & "FROM " & sSaldi & " INNER JOIN ((DecretiIngiuntivi INNER JOIN AnagraficaAvvocati ON DecretiIngiuntivi.CODAVV = AnagraficaAvvocati.CODAVV) "
    qrySQL = qrySQL & "INNER JOIN TribunaliAppartenenza ON DecretiIngiuntivi.CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) ON " & sSaldi & ".Codice = "
    qrySQL = qrySQL & "DecretiIngiuntivi.CODAVV "
    
    
    If Sospeso Then
          qrySQL = qrySQL & "WHERE ((DataEvasionePratica)='') AND ((Annullo)='V') "
     Else
          qrySQL = qrySQL & "WHERE ((DataEvasionePratica)<>'') AND ((Annullo)='V') "
    End If
    
    getQryDecreti = qrySQL
    
End Function


Public Function getQryNotifiche(tipoBimestre As Integer, isUnep As Boolean, tipo As String, Da As String, A As String, provvisorio As String, Optional Sospeso As Boolean) As String
    ' Notifiche
    Dim qrySQL As String
    Dim sSaldi As String
    Dim Tabella As String
    If isUnep Then
      sSaldi = "SaldiUNEP"
      Tabella = "NOTIFICHE_UNEP"
    Else
     sSaldi = "Saldi"
     Tabella = "NOTIFICHE"
    End If
    qrySQL = "SELECT AnagraficaAvvocati.CODAVV, AnagraficaAvvocati.NOME, AnagraficaAvvocati.INDIRI, AnagraficaAvvocati.LOCALI, "
    qrySQL = qrySQL & "AnagraficaAvvocati.PROV, AnagraficaAvvocati.CAP, AnagraficaAvvocati.TELEFCELL, AnagraficaAvvocati.TELEF, NumeroAtto, "
    qrySQL = qrySQL & " ImpNotificheE,'Costo Notifica',ImpSpeseE,DesrSpese, 0,'',0,'',0,'',0,'',ImpCompetenzeE,  "
    If tipo = "Futuro" Then
       If isUnep Then
         qrySQL = qrySQL & "  ImpDepositoE-ImpNotificheE-ImpSpeseE-ImpCompetenzeE,"
        Else
         qrySQL = qrySQL & "  ImpDepositoE-ImpNotificheE-ImpSpeseE-ImpCompetenzeE*" & Str(1 + g_Settings.IVA / 100) & ","
        End If
       Else
       qrySQL = qrySQL & " ImpSaldoE  AS SaldoFinale, "
    End If
    
    qrySQL = qrySQL & " MID(DataPresentazione,7,2) & '/' & MID(DataPresentazione,5,2)& '/' & MID(DataPresentazione,1,4), "
    qrySQL = qrySQL & " MID(DataRestituzione,7,2) & '/' & MID(DataRestituzione,5,2)& '/' & MID(DataRestituzione,1,4), "
    qrySQL = qrySQL & " Crono, "
    
    If Sospeso Then
          qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
          qrySQL = qrySQL & " DataRegistrazione as D , ' ', "

     Else
          qrySQL = qrySQL & " Mid(DataEvasionePratica,7,2) & '/' &  Mid(DataEvasionePratica,5,2) & '/' &  Mid(DataEvasionePratica,1,4), "
          qrySQL = qrySQL & " (MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4)) as D , ' ', "

    End If
    qrySQL = qrySQL & "TribunaliAppartenenza.DescrizioneTribunale, " & sSaldi & ".SaldoTotaleEuro," & sSaldi & ".PROG_Saldi +1,'" + Format(Now, "dd/mm/yyyy") + "',"
    qrySQL = qrySQL & "'E','" & provvisorio & "','" & Da & "','" & A & "', Parte1, Parte2, 'Notifiche', ImpDepositoE,AnagraficaAvvocati.NumOrdinamento,Left(Localita1,18), Note    "
    If isUnep Then
      qrySQL = qrySQL & ", 0 as Quota "
      'qrySQL = qrySQL & "," & FixDouble(IIf(tipoBimestre = 2, g_Settings.QuotaSoci, g_Settings.QuotaSoci / 2)) & " as Quota "
    End If
    qrySQL = qrySQL & "FROM " & sSaldi & " INNER JOIN ((" & Tabella & " INNER JOIN AnagraficaAvvocati ON " & Tabella & ".CODAVV = AnagraficaAvvocati.CODAVV) INNER JOIN TribunaliAppartenenza ON "
    qrySQL = qrySQL & "" & Tabella & ".CodTribunaleApp = TribunaliAppartenenza.CodiceTribunale) ON " & sSaldi & ".Codice = " & Tabella & ".CODAVV "

    
    If Sospeso Then
              qrySQL = qrySQL & "WHERE ((" & Tabella & ".DataEvasionePratica)='') AND ((" & Tabella & ".Annullo)='V') "
     Else
              qrySQL = qrySQL & "WHERE ((" & Tabella & ".DataEvasionePratica)<>'') AND ((" & Tabella & ".Annullo)='V') "
    End If

    getQryNotifiche = qrySQL
    
End Function


Public Function getQrySfratti(tipoBimestre As Integer, isUnep As Boolean, tipo As String, Da As String, A As String, provvisorio As String, Optional Sospeso As Boolean) As String
    ' Sfratti
    Dim qrySQL As String
     Dim sSaldi As String
    Dim Tabella As String
    If isUnep Then
      sSaldi = "SaldiUNEP"
      Tabella = "SFRATTI_UNEP"
    Else
      sSaldi = "Saldi"
      Tabella = "SFRATTI"
    End If
    qrySQL = "SELECT AnagraficaAvvocati.CODAVV, AnagraficaAvvocati.NOME, AnagraficaAvvocati.INDIRI, AnagraficaAvvocati.LOCALI, "
    qrySQL = qrySQL & "AnagraficaAvvocati.PROV, AnagraficaAvvocati.CAP, AnagraficaAvvocati.TELEFCELL, AnagraficaAvvocati.TELEF, NumeroAtto, "
    qrySQL = qrySQL & " ImpSpeseE,'Costo Effettivo',ImpVarieE,DesrSpese, 0,'',0,'',0,'',0,'',ImpCompetenzeE,  "
    If tipo = "Futuro" Then
        If isUnep Then
         qrySQL = qrySQL & "  ImpDepositoE-ImpSpeseE-ImpVarieE-ImpCompetenzeE,"
        Else
         qrySQL = qrySQL & "  ImpDepositoE-ImpSpeseE-ImpVarieE-ImpCompetenzeE*" & Str(1 + g_Settings.IVA / 100) & ","
        End If
       Else
        qrySQL = qrySQL & " ImpSaldoE  AS SaldoFinale, "
    End If
    qrySQL = qrySQL & " MID(DataPresentazione,7,2) & '/' & MID(DataPresentazione,5,2)& '/' & MID(DataPresentazione,1,4), "
    qrySQL = qrySQL & " MID(DataRestituzione,7,2) & '/' & MID(DataRestituzione,5,2)& '/' & MID(DataRestituzione,1,4), "
    qrySQL = qrySQL & " Crono, "
    If Sospeso Then
          qrySQL = qrySQL & " MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4), "
          qrySQL = qrySQL & " DataRegistrazione as D , ' ', "

     Else
          qrySQL = qrySQL & " Mid(DataEvasionePratica,7,2) & '/' &  Mid(DataEvasionePratica,5,2) & '/' &  Mid(DataEvasionePratica,1,4), "
          qrySQL = qrySQL & " (MID(DataRegistrazione,7,2) & '/' & MID(DataRegistrazione,5,2)& '/' & MID(DataRegistrazione,1,4)) as D , ' ', "

    End If
    qrySQL = qrySQL & "TribunaliAppartenenza.DescrizioneTribunale, " & sSaldi & ".SaldoTotaleEuro," & sSaldi & ".PROG_Saldi +1,'" + Format(Now, "dd/mm/yyyy") + "',"
    qrySQL = qrySQL & "'E','" & provvisorio & "','" & Da & "','" & A & "', Parte1, Parte2, 'Sfratti/Pignoramenti', ImpDepositoE,AnagraficaAvvocati.NumOrdinamento,Left(Localita1,35)    "
    If isUnep Then
      qrySQL = qrySQL & ", 0 as Quota "
      'qrySQL = qrySQL & "," & FixDouble(IIf(tipoBimestre = 2, g_Settings.QuotaSoci, g_Settings.QuotaSoci / 2)) & " as Quota "
    End If
    qrySQL = qrySQL & " FROM " & sSaldi & " INNER JOIN ((" & Tabella & " INNER JOIN AnagraficaAvvocati ON " & Tabella & ".CODAVV = "
    qrySQL = qrySQL & " AnagraficaAvvocati.CODAVV) INNER JOIN TribunaliAppartenenza ON " & Tabella & ".CodTribunaleApp "
    qrySQL = qrySQL & " = TribunaliAppartenenza.CodiceTribunale) ON " & sSaldi & ".Codice = " & Tabella & ".CODAVV "
    
    If Sospeso Then
              qrySQL = qrySQL & " WHERE (((" & Tabella & ".DataEvasionePratica)='') AND ((" & Tabella & ".Annullo)='V'))"
     Else
              qrySQL = qrySQL & " WHERE (((" & Tabella & ".DataEvasionePratica)<>'') AND ((" & Tabella & ".Annullo)='V'))"
    End If
    
    
    
    getQrySfratti = qrySQL
End Function


Public Sub update_EstConto_Adempimenti(Tabella As String, qrySQL As String)
    Dim sqlUpdate As String
    
    sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
                "DEPOSITO,SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE, DATA_EVASIONE,DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA,NumOrdinamento) " & _
                qrySQL
    g_Settings.DBConnection.Execute sqlUpdate
End Sub





Public Sub update_EstConto_Decreti(Tabella As String, qrySQL As String)
    Dim sqlUpdate As String
    
    sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
                "SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE, DATA_EVASIONE,DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA," & _
                "QtaCopie,QtaFotocopie,QtaMarche,QtaDirittiCancelleria,Deposito,NumOrdinamento,CodAutorita,Esenzione,FormulaEsec,NumeroIngiunzione,NumeroRuolo) " & _
                qrySQL
    g_Settings.DBConnection.Execute sqlUpdate
End Sub


Public Sub update_EstConto_Notifiche(isUnep As Boolean, Tabella As String, qrySQL As String)
  Dim sqlUpdate As String
    If Tabella = "PrtEstrattoConto" Or Tabella = "PrtEstrattoContoUNEP" Then
        sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
                "SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE,Crono,DATA_EVASIONE, DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA," & _
                "Deposito,NumOrdinamento,Localita1, [Nota] "
                
          
        Else
         sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
                "SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE,Crono,DATA_EVASIONE, DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA," & _
                "Deposito,NumOrdinamento,Localita1, [Note] "
                
               
     End If
     If isUnep Then
        sqlUpdate = sqlUpdate & ", Quota "
     End If
     sqlUpdate = sqlUpdate & "  ) " & qrySQL
    g_Settings.DBConnection.Execute sqlUpdate
End Sub


Public Sub update_EstConto_Sfratti(isUnep As Boolean, Tabella As String, qrySQL As String)
Dim sqlUpdate As String
    sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
                "SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE,Crono, DATA_EVASIONE,DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA," & _
                "Deposito,NumOrdinamento,Localita1"
                
     If isUnep Then
        sqlUpdate = sqlUpdate & ", Quota "
     End If
     sqlUpdate = sqlUpdate & "  ) " & qrySQL
              
    g_Settings.DBConnection.Execute sqlUpdate
End Sub

Public Function getPrecChiusura() As Date
 getPrecChiusura = GetADOValue("Date_EstrattiConto", "DATA_ULTIMO_ESTCONTO", "1=1", g_Settings.DBConnection)
End Function
Public Function getNewNumFattura() As Integer
 getNewNumFattura = GetADOValue("StoricoFatture", "Max(NumeroFattura)", "Left(DataFattura,4)='" & year(Now) & "'", g_Settings.DBConnection, True) + 1
 
End Function
Public Sub AggiungiAvvocatiSenzaOperazioni(data1 As String, data2 As String, codAvv As String, Optional importo As Double)
Dim qrySQL As String
Dim qryApp As String
Dim qryDelete As String
Dim qry1, qry2, qry3 As String

    qry1 = ""
    qry2 = ""
    qry3 = ""
    qryApp = ""
    
    If data1 <> "" Then
       qry1 = " AND ( DataEvasionePratica >= '" & Format(data1, "yyyymmdd") & "')"
    End If
    If data2 <> "" Then
        qry2 = " AND ( DataEvasionePratica <= '" & Format(data2, "yyyymmdd") & "')"
    End If
    
    If codAvv <> "" Then
        qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & codAvv & "')"
    End If
    
    qryApp = qry1 & qry2 & qry3

 qrySQL = "SELECT  CODAVV FROM ANAGRAFICAAVVOCATI " & _
          "WHERE STAT='V' AND NOT (CODAVV LIKE '525%' OR CODAVV LIKE '393%') " & _
          "and CODAVV NOT IN(SELECT CODAVV FROM  SFRATTI_UNEP WHERE 1=1   " & qryApp & ") " & _
          "and CODAVV NOT IN(SELECT CODAVV FROM  NOTIFICHE_UNEP WHERE 1=1   " & qryApp & ") " & _
          "ORDER BY ANAGRAFICAAVVOCATI.NumOrdinamento"
          
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open qrySQL, g_Settings.DBConnection
While Not rs.EOF

    update_EstConto_Notifiche True, "PrtEstrattoContoUNEP", RigaPerAvvocatoSenzaOperazioni(rs(0), data1, data2, importo)
    
  rs.MoveNext
Wend
          
End Sub
Public Sub AggiungiAvvocatiQuota(data1 As String, data2 As String, codAvv As String, Optional importo As Double)
Dim qrySQL As String
Dim qryApp As String
Dim qryDelete As String
Dim qry1, qry2, qry3 As String

    qry1 = ""
    qry2 = ""
    qry3 = ""
    qryApp = ""
    
    If data1 <> "" Then
       qry1 = " AND ( DataEvasionePratica >= '" & Format(data1, "yyyymmdd") & "')"
    End If
    If data2 <> "" Then
        qry2 = " AND ( DataEvasionePratica <= '" & Format(data2, "yyyymmdd") & "')"
    End If
    
    If codAvv <> "" Then
        qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & codAvv & "')"
    End If
    
    qryApp = qry1 & qry2 & qry3

 qrySQL = "SELECT  CODAVV FROM ANAGRAFICAAVVOCATI " & _
          "WHERE STAT='V' AND NOT (CODAVV LIKE '525%' OR CODAVV LIKE '393%') " & _
          "ORDER BY ANAGRAFICAAVVOCATI.NumOrdinamento"
          
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open qrySQL, g_Settings.DBConnection
While Not rs.EOF

    update_EstConto_Notifiche True, "PrtEstrattoContoUNEP", RigaPerAvvocatoSenzaOperazioni(rs(0), data1, data2, importo)
    
  rs.MoveNext
Wend
          
End Sub
Public Function RigaPerAvvocatoSenzaOperazioni(codAvv As String, data1 As String, data2 As String, Optional importo As Double) As String
Dim qrySQL As String
'        sqlUpdate = "INSERT INTO " & Tabella & " (CodAvv,Nome,INDIRI,LOCALI,PROV,CAP,TELEFCELL,TELEF,CRONOLOGICO," & _
'                "SPESE1,DESCR_SPESE1,SPESE2,DESCR_SPESE2,SPESE3,DESCR_SPESE3,SPESE4,DESCR_SPESE4," & _
'                "SPESE5,DESCR_SPESE5,SPESE6,DESCR_SPESE6,COMPETENZE,SALDO,DATA_PRESENTAZIONE, DATA_RESTITUZIONE,Crono,DATA_EVASIONE, DATARegistrazione,AttivitaRichiesta,DESCR_TRIBUNALE," & _
'                "SALDO_PRECEDENTE,NUM_EST_CONTO,DATA_EST_CONTO,VALUTA,PROVVISORIO,DATA_INIZIO,DATA_FINE,Parte1,Parte2,DESCR_ATTIVITA," & _
'                "Deposito,NumOrdinamento,Localita1, [Nota] "

  qrySQL = "SELECT AnagraficaAvvocati.CODAVV, AnagraficaAvvocati.NOME, AnagraficaAvvocati.INDIRI, AnagraficaAvvocati.LOCALI, "
    qrySQL = qrySQL & "AnagraficaAvvocati.PROV, AnagraficaAvvocati.CAP, AnagraficaAvvocati.TELEFCELL, AnagraficaAvvocati.TELEF, 0, "
    qrySQL = qrySQL & " 0,'Costo Notifica',0,'', 0,'',0,'',0,'',0,'',0, 0,'','','','','','', "
    qrySQL = qrySQL & "'', 0,0,'" + Format(Now, "dd/mm/yyyy") + "',"
    qrySQL = qrySQL & "'E','','" & data1 & "','" & data2 & "', '', '', 'Notifiche', 0,AnagraficaAvvocati.NumOrdinamento,'', ''," & FixDouble(importo)
    qrySQL = qrySQL & " FROM AnagraficaAvvocati WHERE CODAVV='" & codAvv & "'"
    
    RigaPerAvvocatoSenzaOperazioni = qrySQL
End Function

Public Sub Riempi_PRT_EstrattoContoX(data1 As String, data2 As String, codAvv As String, _
                                     adempimenti As Integer, Notifiche As Integer, decreti As Integer, sfratti As Integer, _
                                     provvisorio As String, isUnep As Boolean, tipoBimestre As Integer)

Dim qrySQL As String
Dim qryApp As String
Dim qryDelete As String
Dim qry1, qry2, qry3 As String
Dim NumErrori As Integer
    
' Valuta Corrente


On Error GoTo Riempi_PRT_EstrattoConto
    
    qry1 = ""
    qry2 = ""
    qry3 = ""
    qryApp = ""
    
    If data1 <> "" Then
       qry1 = " AND ( DataEvasionePratica >= '" & Format(data1, "yyyymmdd") & "')"
    End If
    If data2 <> "" Then
        qry2 = " AND ( DataEvasionePratica <= '" & Format(data2, "yyyymmdd") & "')"
    End If
    
    If codAvv <> "" Then
        qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & codAvv & "')"
    End If
    
    qryApp = qry1 & qry2 & qry3
    
    OpenProgress ("Attendere... Preparazione Stampa!")
    
    
    'Reset PrtEstrattoConto
    If isUnep Then
      qryDelete = "DELETE FROM PrtEstrattoContoUNEP;"
    Else
      qryDelete = "DELETE FROM PrtEstrattoConto;"
    End If
    
    g_Settings.DBConnection.Execute qryDelete
    
    If adempimenti = 1 Then
            'Inizio Adempimenti
            qrySQL = getQryAdempimenti(isUnep, "", data1, data2, provvisorio) & qryApp & "  ORDER BY ADEMPI.NumOrdinamento"
            update_EstConto_Adempimenti "PrtEstrattoConto", qrySQL
            UpdateProgress 25, "Adempimenti"
            'Fine Adempimenti
    End If
    
    If sfratti = 1 Then
        'Inizio Sfratti
        If isUnep Then
          qrySQL = getQrySfratti(tipoBimestre, True, "", data1, data2, provvisorio) & qryApp & "  ORDER BY SFRATTI_UNEP.NumOrdinamento"
          update_EstConto_Sfratti True, "PrtEstrattoContoUNEP", qrySQL
        Else
          qrySQL = getQrySfratti(0, False, "", data1, data2, provvisorio) & qryApp & "  ORDER BY SFRATTI.NumOrdinamento"
          update_EstConto_Sfratti False, "PrtEstrattoConto", qrySQL
        End If
    
        
        UpdateProgress 50, "Stratti"

        ' Fine Sfratti
    End If
    
    If Notifiche = 1 Then
            'Inizio Notifiche
        If isUnep Then
          qrySQL = getQryNotifiche(tipoBimestre, True, "", data1, data2, provvisorio) & qryApp & "  ORDER BY Notifiche_UNEP.NumOrdinamento"
          update_EstConto_Notifiche True, "PrtEstrattoContoUNEP", qrySQL
        Else
          qrySQL = getQryNotifiche(0, False, "", data1, data2, provvisorio) & qryApp & "  ORDER BY Notifiche.NumOrdinamento"
          update_EstConto_Notifiche False, "PrtEstrattoConto", qrySQL

        End If
             
             UpdateProgress 75, "Notifiche"
           'Fine Notifiche
    End If
       
    If decreti = 1 Then
        'Inizio Decreti
        qrySQL = getQryDecreti(isUnep, "", data1, data2, provvisorio) & qryApp & " ORDER BY DecretiIngiuntivi.NumOrdinamento"
        update_EstConto_Decreti "PrtEstrattoConto", qrySQL
        UpdateProgress 100, "Stampa in corso..."
        'Fine Decreti
    End If
    
    CloseProgress

Exit Sub

Riempi_PRT_EstrattoConto:
   
        MsgBox "Attenzione errore in stampa Estratto Conto!" & Chr(10) & err & " - " & Error(err), vbCritical, "Attenzione"
   
Resume Next
End Sub


Public Sub Riempi_PRT_Sospesi(data1 As String, data2 As String, codAvv As String, _
                              codTribunale As String, codAttività As String, isUnep As Boolean, _
                              orderByData As Boolean, tipoBimestre As Integer)

Dim qrySQL As String
Dim qryDelete As String
Dim qryAppAd As String
Dim qry1, qry2, qry3, qryTrib As String
Dim qryAppSfr As String
Dim currentTable As String
Dim qryAppDec As String

Dim qryAppNot As String


Dim NumErrori As Integer

    
On Error GoTo Riempi_PRT_Sospesi
    
    OpenProgress ("Attendere... Preparazione Stampa!")
    
    'Reset PrtEstrattoConto
Dim table As String

Dim unepWhere As String
    If isUnep Then
      table = "PrtSospesiUNEP"
      
    Else
      table = "PrtSospesi"
      
    End If
    
    qryDelete = "DELETE  * FROM " & table
    g_Settings.DBConnection.Execute qryDelete

        If codTribunale <> "NULL" Then
            qryTrib = " AND ( CodTribunaleApp = '" & codTribunale & "')"
        End If
        If data1 <> "" Then
            qry1 = " AND ( DataRegistrazione >= '" & Format(data1, "yyyymmdd") & "')"
        End If
        If data2 <> "" Then
            qry2 = " AND ( DataRegistrazione <= '" & Format(data2, "yyyymmdd") & "')"
        End If
        If codAvv <> "" Then
            qry3 = " AND ( AnagraficaAvvocati.CODAVV = '" & codAvv & "')"
        End If
     
  
   If Not isUnep Then
        'Inizio Adempimenti
        If codAttività = "NULL" Or codAttività = "A" Then
            qryAppAd = qryTrib & qry1 & qry2 & qry3
            If orderByData Then
             qrySQL = getQryAdempimenti(isUnep, "Attuale", data1, data2, "S", True) & qryAppAd & " ORDER BY ADEMPI.DataRegistrazione"
            Else
             qrySQL = getQryAdempimenti(isUnep, "Attuale", data1, data2, "S", True) & qryAppAd & " ORDER BY ADEMPI.NumOrdinamento"
            End If
            update_EstConto_Adempimenti table, qrySQL
            UpdateProgress (5)
        End If
        'Fine Adempimenti
    End If
    
    
    'Inizio Sfratti
    
    If codAttività = "NULL" Or codAttività = "S" Then
        currentTable = "SFRATTI"
        If isUnep Then currentTable = currentTable & "_UNEP"
        qryAppSfr = qryTrib & qry1 & qry2 & qry3
        If orderByData Then
         qrySQL = getQrySfratti(tipoBimestre, isUnep, "Attuale", data1, data1, "S", True) & qryAppSfr & " ORDER BY " & currentTable & ".DataRegistrazione"
        Else
         qrySQL = getQrySfratti(tipoBimestre, isUnep, "Attuale", data1, data1, "S", True) & qryAppSfr & " ORDER BY " & currentTable & ".NumOrdinamento"
        End If
        'qrySQL = qrySQL & qryAppSfr
        update_EstConto_Sfratti isUnep, table, qrySQL
        UpdateProgress (30)
    End If
    ' Fine Sfratti
    
    
    
    'Inizio Notifiche
    If codAttività = "NULL" Or codAttività = "N" Then
        currentTable = "Notifiche"
        If isUnep Then currentTable = currentTable & "_UNEP"
        qryAppNot = qryTrib & qry1 & qry2 & qry3
        If orderByData Then
           qrySQL = getQryNotifiche(tipoBimestre, isUnep, "Attuale", data1, data1, "S", True) & qryAppNot & " ORDER BY " & currentTable & ".DataRegistrazione"
        Else
            qrySQL = getQryNotifiche(tipoBimestre, isUnep, "Attuale", data1, data1, "S", True) & qryAppNot & " ORDER BY " & currentTable & ".NumOrdinamento"
        End If
        update_EstConto_Notifiche isUnep, table, qrySQL
        UpdateProgress (50)
    End If
    'Fine Notifiche


    
    If Not isUnep Then
        'Inizio Decreti
        If codAttività = "NULL" Or codAttività = "D" Then
            qryAppDec = qryTrib & qry1 & qry2 & qry3
            If orderByData Then
                qrySQL = getQryDecreti(isUnep, "Attuale", data1, data1, "S", True) & qryAppDec & " ORDER BY DecretiIngiuntivi.DataRegistrazione"
            Else
                qrySQL = getQryDecreti(isUnep, "Attuale", data1, data1, "S", True) & qryAppDec & " ORDER BY DecretiIngiuntivi.NumOrdinamento"
            End If
            update_EstConto_Decreti table, qrySQL
            UpdateProgress (80)
        End If
    'Fine Decreti
    End If
    CloseProgress

Exit Sub

Riempi_PRT_Sospesi:
  MsgBox "Attenzione errore in stampa Sospesi!" & Chr(10) & err & " - " & Error(err), vbCritical, "Attenzione"


End Sub


