Attribute VB_Name = "libRST"
Option Explicit

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Function SafeMakeDir(D As String)
 If Dir(D, vbDirectory) = "" Then
     MkDir (D)
 End If
End Function
Public Function SafeKill(f As String)
On Error GoTo esci
If Dir(f) <> "" Then Kill f
esci:
End Function

Public Sub formattaSaldo(lbl As Label, x As Double)
lbl.Caption = Format(x, "##,##0.00")

If x < 0 Then lbl.ForeColor = RGB(255, 0, 0) Else lbl.ForeColor = RGB(0, 0, 255)

End Sub
Public Function getPath(s As String) As String
 Dim p As Integer
 p = InStrRev(s, "\")
 getPath = Left(s, p)
End Function
Public Function GetFileNameWithoutExtension(s As String) As String
 Dim sFile As String
 sFile = GetFile(s)
 Dim L As Integer, I As Integer
 L = Len(sFile)
 For I = L To 1 Step -1
  If Mid(sFile, I, 1) = "." Then Exit For
 Next I
 GetFileNameWithoutExtension = Left(sFile, I - 1)
 
End Function
Public Function GetFile(s As String) As String
 Dim L As Integer, I As Integer
 L = Len(s)
 For I = L To 1 Step -1
  If Mid(s, I, 1) = "\" Then Exit For
 Next I
 GetFile = Mid(s, I + 1)
 

End Function
Public Sub RiempiTestata(frm As Form, rs As ADODB.Recordset)
    frm.LblCodiceA.Caption = rs!codAvv
    frm.LblDescrCodAvv.Caption = rs!nome
    frm.TxtCodiceAvvocato = rs!codAvv
End Sub
Public Function ExistMDBTable(CN As ADODB.Connection, table As String)

On Local Error Resume Next
CN.Execute ("SELECT * FROM " & table)
ExistMDBTable = (err.Number = 0)
err.Clear

End Function
Public Function ExistMDBField(CN As ADODB.Connection, table As String, Field As String)
Dim rsX As ADODB.Recordset
On Local Error Resume Next
err.Clear
CN.Execute ("SELECT " & Field & " FROM " & table)
ExistMDBField = (err.Number = 0)
err.Clear
rsX.Close
End Function
Public Sub SelectItemInTDBCombo(cmb As TDBCombo, value As Long)
Dim r
           r = cmb.Columns(1).Find(value, dblSeekEQ, True)
           If Not IsNull(r) Then cmb.Bookmark = r
           cmb.BoundText = cmb.Columns(0).value
End Sub
Public Sub Caricacampi(frm As Form, rs As ADODB.Recordset, Optional noLabel As Boolean)

Dim c As Control
Dim I As Integer
Dim r
Dim x
  For Each c In frm.Controls
 
        If TypeOf c Is Label And noLabel = False Then
           If c.DataField <> "" Then
              c = IIf(IsNull(rs(c.DataField)), "", rs(c.DataField))
           End If
        End If
 
    
        If TypeOf c Is TextBox Then
           If c.DataField <> "" Then
              c = IIf(IsNull(rs(c.DataField)), "", rs(c.DataField))
           End If
        End If
        If TypeOf c Is TDBNumber Then
           If c.DataField <> "" Then
              c = IIf(IsNull(rs(c.DataField)), 0, rs(c.DataField))
           End If
        End If
         If TypeOf c Is TDBDate Then
           If IsDate(rs(c.DataField)) Then
              c = rs(c.DataField)
           End If
        End If
'        If TypeOf c Is ComboBox Then
'          For i = 1 To c.ListCount
'            If frm.colTrib(i) = rs(c.DataField) Then Exit For
'          Next i
'          If (i - 1) < c.ListCount Then c.ListIndex = i - 1
'        End If
        If TypeOf c Is TDBCombo Then
           r = c.Columns(1).Find(rs(c.DataField).value, dblSeekEQ, True)
           If Not IsNull(r) Then c.Bookmark = r
           c.BoundText = c.Columns(0).value
        End If
        
        If TypeOf c Is CheckBox Then
           Select Case c.DataField
             Case "Annullo", "STAT"
               c.value = IIf(IsNull(rs(c.DataField)), 0, -(rs(c.DataField) = "A"))
             Case "CheckVisual"
               c.value = IIf(IsNull(rs(c.DataField)), 0, -(rs(c.DataField) = "X"))
             Case "IsUNEP"
               c.value = IIf(IsNull(rs(c.DataField)), 0, IIf(rs(c.DataField), 1, 0))
             Case Else
               c.value = IIf(IsNull(rs(c.DataField)), 0, -(rs(c.DataField) = "S"))
           End Select
           
       End If
 
  Next
  
End Sub
Public Sub Dimensiona(D As String, f As Form)
Select Case D
 Case "BIG"
  f.Move 3000, 200, 6500, 5000
 Case "SMALL"
  f.Move 3000, 200, 6500, 2200
 End Select
End Sub
Public Sub CloseAllForms()
Dim I As Integer

For I = 0 To Forms.count - 1
  If Forms(I).name <> "Atap" And Forms(I).name <> "frmback" Then Unload Forms(I)
Next I
End Sub
Public Function FindForm(ByVal form_name As String) As Boolean
    Dim I As Integer
   
    ' Search the loaded forms.
    For I = 0 To Forms.count - 1
        If Forms(I).name = form_name Then
         
            FindForm = True
            Exit For
        End If
    Next I
End Function
Public Function ExistADORecord(SQL As String, Conn As ADODB.Connection) As Boolean
 Dim rs As ADODB.Recordset
 Set rs = newAdoRs
 
 rs.Open SQL, Conn
 ExistADORecord = (rs.RecordCount > 0)
 rs.Close
 
End Function
Public Function GetADOValue(table As String, Field As String, Where As String, Conn As ADODB.Connection, Optional isNumber As Boolean)
 Dim rs As ADODB.Recordset
 Dim SQL As String
 Dim val
 Set rs = newAdoRs
 SQL = "SELECT " & Field & " FROM " & table & " WHERE " & Where
 rs.Open SQL, Conn
 If rs.EOF Then
   val = ""
  Else
   val = IIf(IsNull(rs(0)), "", rs(0))
 End If
 If isNumber And val = "" Then GetADOValue = 0 Else GetADOValue = val
 rs.Close
 Set rs = Nothing
End Function
Public Function GetADORecordset(table As String, Fields As String, Where As String, Conn As ADODB.Connection) As ADODB.Recordset
 Dim rs As ADODB.Recordset
 Dim SQL As String
 Set rs = newAdoRs
 SQL = "SELECT " & Fields & " FROM " & table & " WHERE " & Where
 rs.Open SQL, Conn
 If rs.EOF Then
   Set GetADORecordset = Nothing
  Else
   Set GetADORecordset = rs
 End If
End Function

Public Function SalvaRecord(frm As Form, tipo As TipoAzione, Tabella As String, Progressivo As Boolean, Optional Where As String, Optional noLabel As Boolean) As Boolean
On Error GoTo fine
Dim SQL As String
Dim c As Control
Dim codTribunale As String
If Progressivo Then codTribunale = frm.cmbTribunale.Columns(1).value
tipo = UCase(tipo)
g_Settings.DBConnection.BeginTrans
If tipo = TipoAzione.Nuovo Then
  
        If Progressivo Then frm.LblNumeroAtto = val(GetADOValue(Tabella, "MAX(" & frm.Tag & ")", "Left(DataRegistrazione,4)='" & Format(frm.txtDataReg, "YYYY") & "' AND CodTribunaleApp='" & codTribunale & "'", g_Settings.DBConnection)) + 1
          SQL = "INSERT INTO " & Tabella & " ("
          For Each c In frm.Controls
            If TypeOf c Is TextBox Or TypeOf c Is TDBNumber Or TypeOf c Is TDBDate Or TypeOf c Is TDBCombo Or TypeOf c Is ComboBox Or TypeOf c Is CheckBox Or (TypeOf c Is Label And noLabel = False) Then
                If c.DataField <> "" And c.Tag <> "XXX" Then
                    SQL = SQL & "[" & c.DataField & "],"
                End If
            End If
              
          Next c
           SQL = Left(SQL, Len(SQL) - 1) & ") VALUES ("
         For Each c In frm.Controls
           If TypeOf c Is TextBox Then
             If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & "'" & Replace(c, "'", "''") & "',"
           End If
           
            If TypeOf c Is TDBNumber Then
             If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & Str(c) & ","
           End If
          
           If TypeOf c Is TDBDate Then
             If c.DataField <> "" Then SQL = SQL & "'" & Format(c, "YYYYMMDD") & "',"
           End If
           
           If TypeOf c Is TDBCombo Then
              If c.DataField <> "" Then SQL = SQL & "'" & c.Columns(1).value & "',"
           End If
           
             
           If TypeOf c Is CheckBox Then
              Select Case c.DataField
                Case "Annullo", "STAT"
                  SQL = SQL & "'" & IIf(c.value = 1, "A", "V") & "',"
                Case "CheckVisual"
                  SQL = SQL & "'" & IIf(c.value = 1, "X", "") & "',"
                Case "IsUNEP"
                  SQL = SQL & " " & IIf(c.value = 1, "true", "false") & ","
                Case Else
                  SQL = SQL & "'" & IIf(c.value = 1, "S", "N") & "',"
              End Select
           End If
           
           If TypeOf c Is Label And noLabel = False Then
              If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & "'" & c & "',"
           End If
           
         Next
           SQL = Left(SQL, Len(SQL) - 1) & ")"
Else
             SQL = "UPDATE " & Tabella & " SET "
            
             
            For Each c In frm.Controls
              If TypeOf c Is TextBox Then
                If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & "[" & c.DataField & "]='" & Replace(c, "'", "''") & "',"
              End If
              If TypeOf c Is TDBNumber Then
                If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & "[" & c.DataField & "]=" & Str(c.value) & ","
              End If
              If TypeOf c Is TDBDate Then
                If c.DataField <> "" Then SQL = SQL & "[" & c.DataField & "]='" & Format(c, "YYYYMMDD") & "',"
              End If
              
              If TypeOf c Is TDBCombo Then
                 If c.DataField <> "" Then
                       If IsEmpty(c.Columns(1).value) Then
                         SQL = SQL & "[" & c.DataField & "]='" & c.Tag & "',"
                       Else
                         SQL = SQL & "[" & c.DataField & "]='" & c.Columns(1).value & "',"
                       End If
                 End If
                  
                 
              End If
              
              If TypeOf c Is CheckBox Then
                 Select Case c.DataField
                   Case "Annullo", "STAT"
                     SQL = SQL & "[" & c.DataField & "]='" & IIf(c.value = 1, "A", "V") & "',"
                   Case "CheckVisual"
                     SQL = SQL & "[" & c.DataField & "]='" & IIf(c.value = 1, "X", "") & "',"
                   Case "IsUNEP"
                     SQL = SQL & "[" & c.DataField & "]=" & IIf(c.value = 1, "true", "false") & ","
                   Case Else
                    SQL = SQL & "[" & c.DataField & "]='" & IIf(c.value = 1, "S", "N") & "',"
                 End Select
                 
              End If
              
              If TypeOf c Is Label And noLabel = False Then
                 If c.DataField <> "" And c.Tag <> "XXX" Then SQL = SQL & "[" & c.DataField & "]='" & c & "',"
              End If
              
            Next
              SQL = Left(SQL, Len(SQL) - 1) & " WHERE " & Where


    End If
    'Ultimo controllo sul lock delle tabelle
    If tipo = TipoAzione.Nuovo Then
      If IsTableLocked(Tabella) Then
        SQL = ""
        err.Raise 1, Tabella, "Tabella " & Tabella & " bloccata da un altro utente."
      End If
    Else
      If IsRecordLocked(Where, Tabella) Then
        SQL = ""
        err.Raise 1, Tabella, "Record della " & Tabella & " bloccato da un altro utente."
      End If
    End If
    
    
    g_Settings.DBConnection.Execute SQL
    
    SalvaRecord = True
    g_Settings.DBConnection.CommitTrans
    Exit Function
fine:
   If err.Number = -2147467259 Then 'Indice duplicato
     If Progressivo Then
        frm.LblNumeroAtto = val(GetADOValue(Tabella, "MAX(" & frm.Tag & ")", "Left(DataRegistrazione,4)='" & Format(frm.txtDataReg, "YYYY") & "' AND CodTribunaleApp='" & codTribunale & "'", g_Settings.DBConnection)) + 1
        MsgBox "Indice Duplicato!! " & vbCrLf & "Il sistema ha variato automaticamante il numero progressivo." & vbCrLf & "Riprovare premendo di nuovo su <Modifica>", vbOKOnly + vbInformation, "Atap"
       Else
        MsgBox "Indice Duplicato: la tabella contiene valori che non possono essere duplicati!! " & vbCrLf & "Ad esempio il codice deve essere univoco. " & vbCrLf & "Apportare le modifiche necessarie e riprovare.", vbOKOnly + vbInformation
      End If
      'MsgBox err.Description & vbCrLf & err.Source & vbCrLf & SQL
    Else
     MsgBox err.Description & vbCrLf & err.Source & vbCrLf & SQL
   End If
   SalvaRecord = False
   g_Settings.DBConnection.RollbackTrans

End Function
Public Sub TipoMaschera(frm As Form, tipo As TipoAzione)
Dim c As Frame
'For Each c In frm.fraMaschera
'   c.Visible = (tipo = TipoAzione.Nuovo)
'Next
frm.CmdSalva.Caption = IIf(tipo = TipoAzione.Nuovo, "Salva", "Modifica")
frm.CmdSalva.Visible = (tipo <> TipoAzione.Vuoto)
Call Ridimensiona(frm, IIf(tipo <> TipoAzione.Vuoto, "big", "small"))
Select Case tipo
  Case TipoAzione.Vuoto
        Call PulisciCampi(frm)
        Call PulisciTestata(frm)
        frm.CmdRicerca.Enabled = True
        frm.CmdRicercaAnag.Enabled = True
  Case TipoAzione.Nuovo
        frm.txtDataReg.value = Date
        frm.CmdRicerca.Enabled = False
        frm.CmdRicercaAnag.Enabled = False
   Case Else
        frm.CmdRicerca.Enabled = False
        frm.CmdRicercaAnag.Enabled = False
       
End Select


End Sub
Public Function SalvaTutto(frm As Form, Tabella As String, sWhere As String, Progressivo As Boolean, Optional isUnep As Boolean) As Boolean
Dim e As String
Dim XXX As Boolean
Dim Response
e = ControllaInput(frm, isUnep)
If e = "" Then
            If frm.Azione = TipoAzione.Nuovo Then
              'voce nuova
                Response = MsgBox("Vuoi salvare i dati inseriti?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
                If Response = vbYes Then    ' User chose Yes.
                  XXX = SalvaRecord(frm, frm.Azione, Tabella, Progressivo)
                  
                End If
             Else
              'modifica
                Response = MsgBox("Vuoi salvare le modifiche effettuate?", vbYesNo + vbInformation + vbDefaultButton2, "Attenzione")
                If Response = vbYes Then    ' User chose Yes.
                  XXX = SalvaRecord(frm, frm.Azione, Tabella, Progressivo, sWhere)
                  
                End If
            End If
            
            
                If XXX Then Call TipoMaschera(frm, TipoAzione.Vuoto)
                   

 
   Else
    MsgBox e, vbCritical + vbOKOnly, "Atap"
   
   End If
SalvaTutto = XXX
End Function
Public Function ControllaInput(frm As Form, Optional isUnep As Boolean)
Dim e As String
Dim c As Control
For Each c In frm.Controls
  If (isUnep And Left(c.Tag, 10) = "necessUNEP") Or Left(c.Tag, 10) = "necessario" Then
     
       If TypeOf c Is TextBox Then
            If Trim(c.Text) = "" Then e = e & Mid(c.Tag, 12) & vbCrLf
       ElseIf TypeOf c Is TDBCombo Then
            If Trim(c.Text) = "" Then e = e & Mid(c.Tag, 12) & vbCrLf
       ElseIf TypeOf c Is TDBNumber Then
            If Trim(c.Text) = "" Then e = e & Mid(c.Tag, 12) & vbCrLf
       ElseIf TypeOf c Is TDBDate Then
            If Not IsDate(c.Text) Then e = e & Mid(c.Tag, 12) & vbCrLf
       End If
   
    
 End If
Next
 If e <> "" Then e = "Mancano i seguenti dati:" & vbCrLf & e
 ControllaInput = e
End Function
Public Function ControllaTribunale(cod As String) As String
  ControllaTribunale = ""
  If Not GetADORecordset("ADEMPI", "CodTribunaleApp", "CodTribunaleApp='" & cod & "'", g_Settings.DBConnection) Is Nothing Then
      ControllaTribunale = ControllaTribunale & " - Adempimenti;" & vbCrLf
      
  End If
  
  If Not GetADORecordset("DecretiIngiuntivi", "CodTribunaleApp", "CodTribunaleApp='" & cod & "'", g_Settings.DBConnection) Is Nothing Then
     ControllaTribunale = ControllaTribunale & " - Decreti Ingiuntivi;" & vbCrLf
      Exit Function
  End If
  
  If Not GetADORecordset("Notifiche", "CodTribunaleApp", "CodTribunaleApp='" & cod & "'", g_Settings.DBConnection) Is Nothing Then
            ControllaTribunale = ControllaTribunale & " - Notifiche;" & vbCrLf
  End If
  
  If Not GetADORecordset("Sfratti", "CodTribunaleApp", "CodTribunaleApp='" & cod & "'", g_Settings.DBConnection) Is Nothing Then
            ControllaTribunale = ControllaTribunale & " - Sfratti;" & vbCrLf
  End If
  
  If Not GetADORecordset("Anticipi", "CodiceTribunale", "CodiceTribunale='" & cod & "'", g_Settings.DBConnection) Is Nothing Then
            ControllaTribunale = ControllaTribunale & " - Anticipi;" & vbCrLf
  End If
  
End Function

Public Sub AggiornaGriglia(flex As VSFlexGrid, qry As String, Optional btnElimina As CommandButton, Optional btnModifica As CommandButton)
Dim rs As ADODB.Recordset
Set rs = newAdoRs

rs.Open qry, g_Settings.DBConnection
Set flex.DataSource = rs
If Not btnElimina Is Nothing Then
 btnElimina.Enabled = (rs.RecordCount > 0)
End If

If Not btnModifica Is Nothing Then
 btnModifica.Enabled = (rs.RecordCount > 0)
End If
flex.ColWidth(1) = 1000
flex.ColWidth(2) = 3500
End Sub

Public Sub ErrLogFile(f As String, msg, Optional contenuto1, Optional contenuto2, _
                      Optional contenuto3, Optional contenuto4, Optional contenuto5)
Dim fN As Long
fN = FreeFile
   MsgBox err.Description & vbCrLf & err.Source
   
   
 Open app.Path & "\" & f For Append As fN

   
    Write #fN, ">>>" & Now
    Write #fN, "---------------------------------------------------------"
    Write #fN, contenuto1
    Write #fN, contenuto2
    Write #fN, contenuto3
    Write #fN, contenuto4
    Write #fN, contenuto5
    Write #fN, "---------------------------------------------------------"
    Close fN
    
    If err = 3000 Then
        Resume Next
    End If

End Sub
Function CheckCodFiscPIva(Code As String) As Boolean
    Dim A, b, c, DECODE, valori, nro, trasfrom, trasto
    Dim I As Integer
    valori = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    DECODE = "010005070913151719210204182011030608121416102225242301000507091315171921"
    trasfrom = "0123456789"
    trasto = "0246813579"
    nro = 0
    If Len(Trim(Code)) <> 16 And Len(Trim(Code)) <> 11 Then
        'MsgBox ("Lunghezza non corretta")
        MsgBox ("Lunghezza P/IVA o C/FISCALE non corretti!" + vbCrLf), vbQuestion
        Exit Function
    End If
    'controllo codice fiscale
    If Len(Trim(Code)) = 16 Then
      I = 1
      Do While I < 16
        A = Mid(Code, I, 1)
        c = InStr(valori, A)
        If c = 0 Then
          'MsgBox ("Codice Fiscale contiene caratteri non corretti!" + vbCrLf), vbQuestion
          'MsgBox ("Caratteri non corretti")
          Exit Function
        End If
        b = Int(I / 2)
        b = b * 2
        If b = I Then
          nro = nro + c - 1
        Else
          nro = nro + val(Mid(DECODE, c * 2 - 1, 2))
        End If
        I = I + 1
      Loop
      nro = nro - Int(nro / 26) * 26 + 1
      If Mid(Code, 16, 1) = Mid(valori, nro, 1) Then
        CheckCodFiscPIva = True
      Else
        'MsgBox ("Codice Fiscale non corretto!" + vbCrLf), vbQuestion
        'MsgBox ("Codice Fiscale non corretto")

      End If
      Exit Function
    End If
    '
    ' controllo partita iva
    '
    I = 1
    Do While I < 12
        A = Mid(Code, I, 1)
        c = InStr(trasfrom, A)
        If c = 0 Then
            'MsgBox ("Partita IVA contiene caratteri non corretti!" + vbCrLf), vbQuestion
            'MsgBox ("Caratteri non corretti")
            Exit Function
        End If
        b = Int(I / 2)
        b = b * 2
        If b = I Then
            nro = nro + val(Mid(trasto, c, 1))
        Else
            nro = nro + val(Mid(Code, I, 1))
        End If
        I = I + 1
    Loop
    nro = nro - Int(nro / 10) * 10
    If nro = 0 Then
        CheckCodFiscPIva = True
    Else
        'MsgBox ("Partita IVA non corretta!" + vbCrLf), vbQuestion
        
    End If
End Function


Public Function newAdoRs() As ADODB.Recordset
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  rs.CursorType = adOpenForwardOnly
  rs.LockType = adLockPessimistic
  Set newAdoRs = rs
End Function
Public Function LastDay(month As Integer, year As Integer) As Date

LastDay = year & "/" & month & "/" & LastDayOfMonth(month, year)



End Function
Public Function LastDayOfMonth(m As Integer, A As Integer) As Integer
 Dim gg
 Dim g As Integer
 
 gg = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
 
 g = gg(m - 1)
 If m = 2 And Bisestile(A) Then g = g + 1
 
 LastDayOfMonth = g
 

End Function
Public Function Bisestile(A As Integer) As Boolean
  Bisestile = ((A Mod 4 = 0) And (A Mod 100 <> 0)) Or (A Mod 400 = 0)
End Function

