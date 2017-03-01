Attribute VB_Name = "mdlLockManagement"
Option Explicit

Public Function IsRecordLocked(Where As String, table As String) As Boolean
  Dim Owner As String
  Owner = GetADOValue(table, "Locked", Where, g_Settings.DBConnection)
  IsRecordLocked = (Owner <> "NONE" And Owner <> g_Settings.UserLock)
End Function
Public Sub LockRecord(id As Long, table As String)
  g_Settings.DBConnection.Execute "UPDATE " & table & " SET Locked='" & g_Settings.UserLock & "' WHERE IDCod=" & id
End Sub
Public Sub DeLockRecord(id As Long, table As String)
 g_Settings.DBConnection.Execute "UPDATE " & table & " SET Locked='NONE' WHERE IDCod=" & id
End Sub
Public Sub DeLockAllRecord(table As String)
 g_Settings.DBConnection.Execute "UPDATE " & table & " SET Locked='NONE' WHERE Locked='" & g_Settings.UserLock & "'"
End Sub
Public Sub LockTable(table As String)
    g_Settings.DBConnection.Execute "UPDATE LockTable SET Locked='" & g_Settings.UserLock & "' WHERE UCase(TabID)='" & UCase(table) & "'"
 
End Sub
Public Sub DelockTable(table As String)
  g_Settings.DBConnection.Execute "UPDATE LockTable SET Locked='NONE' WHERE UCase(TabID)='" & UCase(table) & "'"
End Sub
Public Function IsTableLocked(table As String) As Boolean
  Dim Owner As String
  Owner = GetADOValue("LockTable", "Locked", "UCase(TabID)='" & UCase(table) & "'", g_Settings.DBConnection)
  IsTableLocked = (Owner <> "" And Owner <> "NONE" And Owner <> g_Settings.UserLock)
End Function
Public Sub DeLockAllTables(Optional onlyUser As Boolean)
Dim sWhere As String
Dim v As String
Dim n As Integer
If onlyUser Then
    v = Left(ComputerName, 15) & Left(UserName, 15)
    n = Len(v)
    
    sWhere = " WHERE Left(Locked," & n & ")='" & v & "'"
End If
g_Settings.DBConnection.Execute "UPDATE LockTable SET " & _
                 "Locked='NONE' " & sWhere
g_Settings.DBConnection.Execute "UPDATE ADEMPI SET " & _
                 "Locked='NONE' " & sWhere
g_Settings.DBConnection.Execute "UPDATE DecretiIngiuntivi SET " & _
                 "Locked='NONE' " & sWhere
g_Settings.DBConnection.Execute "UPDATE Notifiche SET " & _
                 "Locked='NONE' " & sWhere
g_Settings.DBConnection.Execute "UPDATE Sfratti SET " & _
                 "Locked='NONE' " & sWhere
g_Settings.DBConnection.Execute "UPDATE Sfratti_UNEP SET " & _
                 "Locked='NONE' " & sWhere

g_Settings.DBConnection.Execute "UPDATE Notifiche_UNEP SET " & _
                 "Locked='NONE' " & sWhere
                 
g_Settings.DBConnection.Execute "UPDATE Deduzioni_UNEP SET " & _
                 "Locked='NONE' " & sWhere

If Not onlyUser Then MsgBox "Tabelle sbloccate.", vbInformation

                 
End Sub
Public Sub DeLockAllPrtTables()
g_Settings.DBConnection.Execute "UPDATE LockPrt SET " & _
                 "PrtAssegniCircolari=FALSE," & _
                   "PrtEstrattoConto =FALSE," & _
                   "PrtFattProv =FALSE," & _
                   "PrtGiornalieraAdempimenti =FALSE," & _
                   "PrtGiornalieraDecretiIngiuntivi =FALSE," & _
                   "PrtGiornalieraNotifiche =FALSE," & _
                   "PrtGiornalieraNotificheUNEP =FALSE," & _
                   "PrtGiornalieraSfrattiPig =FALSE," & _
                   "PrtGiornalieraSfrattiPigUNEP =FALSE," & _
                   "PrtSaldi =FALSE," & _
                   "PrtAssegniCircolariUNEP=FALSE," & _
                   "PrtEstrattoContoUNEP =FALSE," & _
                   "PrtSospesiUNEP =FALSE," & _
                   "PrtSaldiUNEP =FALSE," & _
                   "PrtSospesi=FALSE;"
MsgBox "Stampe sbloccate.", vbInformation
End Sub
Public Sub LockPrtTable(table As String)
  g_Settings.DBConnection.Execute ("UPDATE LockPrt SET " & table & " = True;")
  
End Sub
Public Sub DelockPrtTable(table As String)
  g_Settings.DBConnection.Execute ("UPDATE LockPrt SET " & table & " = False;")
  
End Sub
Public Function IsPrtTableLocked(table As String)
Dim rs As ADODB.Recordset

Set rs = newAdoRs
rs.Open "SELECT " & table & " FROM LockPrt;", g_Settings.DBConnection

IsPrtTableLocked = rs.Fields(0).value
rs.Close
End Function


