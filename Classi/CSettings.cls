VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Dbname As String
Private m_reportPath As String
Private m_dbPath As String
Private m_LocalTempPath As String
Private m_AutoBackupPath As String
Private m_StoricoPath As String
Private m_StoricoLiquidazioniPath As String
Private m_StoricoECPath As String
Private m_StoricoECAdempiPath As String
Private m_StoricoEC_UNEP As String
Private m_TIMESTAMP As String
Private m_USER_LOCK As String
Private m_Conn As ADODB.Connection
Private m_IVA As Double
Private m_QuotaSoci As Double
Private m_LimiteSaldo As Double
Private m_Banca As String
Private m_IBAN As String

Private m_CurrentDbName As String
Private Const K_NomeInRegistry As String = "AtapDB"



Private Sub Class_Initialize()
    
     Dim frm As frmPathDB
     Dim dbUpdater As CDBUpdater
     Set dbUpdater = New CDBUpdater
     m_TIMESTAMP = Format(Now, "YYMMDDhhmmss")
     m_USER_LOCK = GetUserLock

     
     LeggeRegistroPerDB (K_NomeInRegistry)
     m_reportPath = app.Path & "\Report"
    
    Do
      m_dbPath = getPath(m_Dbname)
      creaCartelle
      'Set m_Conn = ConnettiSQL()
      Set m_Conn = Connetti(m_Dbname)
      If m_Conn Is Nothing Then
        Set frm = New frmPathDB
        frm.Initialize m_Dbname, m_dbPath
        frm.Show vbModal
        If frm.ExitMode = exitOk Then
         m_Dbname = frm.DBName
         ScriviRegistro K_NomeInRegistry, "path", "database", m_Dbname
        End If
     
      End If
    Loop While m_Conn Is Nothing
    'Set GLO_Report = Connetti(app.Path & "\Report.mdb")
    dbUpdater.UpdateDatabase m_Conn
    m_IVA = GetADOValue("Parametri", "IVA", "1=1", Me.DBConnection)
    m_QuotaSoci = GetADOValue("Parametri", "QuotaSoci", "1=1", Me.DBConnection)
    m_LimiteSaldo = GetADOValue("Parametri", "LimiteSaldo", "1=1", Me.DBConnection)
    m_Banca = GetADOValue("Parametri", "Banca", "1=1", Me.DBConnection)
    m_IBAN = GetADOValue("Parametri", "IBAN", "1=1", Me.DBConnection)
End Sub

Public Sub ConnettiDB(sFile As String)
 m_Conn.Close
 Set m_Conn = Connetti(sFile)
 m_CurrentDbName = sFile
End Sub
Public Property Get DBConnection() As ADODB.Connection
 Set DBConnection = m_Conn
End Property
Public Property Get TimeStamp() As String
 TimeStamp = m_TIMESTAMP
End Property
Public Property Get UserLock() As String
 UserLock = m_USER_LOCK
End Property
Public Property Get IVA() As Double
 IVA = m_IVA
End Property
Public Property Let IVA(X As Double)
  m_IVA = X
  Me.DBConnection.Execute "UPDATE Parametri SET IVA=" & Str(m_IVA)
End Property
Public Property Get QuotaSoci() As Double
 QuotaSoci = m_QuotaSoci
End Property
Public Property Let QuotaSoci(X As Double)
  m_QuotaSoci = X
  Me.DBConnection.Execute "UPDATE Parametri SET QuotaSoci=" & Str(m_QuotaSoci)
End Property
Public Property Let LimiteSaldo(X As Double)
  m_LimiteSaldo = X
  Me.DBConnection.Execute "UPDATE Parametri SET LimiteSaldo=" & Str(m_LimiteSaldo)
End Property
Public Property Get LimiteSaldo() As Double
 LimiteSaldo = m_LimiteSaldo
End Property
Public Property Let Banca(X As String)
  m_Banca = X
  Me.DBConnection.Execute "UPDATE Parametri SET Banca='" & Replace(m_Banca, "'", "''") & "'"
End Property
Public Property Get Banca() As String
 Banca = m_Banca
End Property
Public Property Let IBAN(X As String)
  m_IBAN = X
  Me.DBConnection.Execute "UPDATE Parametri SET IBAN='" & Replace(m_IBAN, "'", "''") & "'"
End Property
Public Property Get IBAN() As String
 IBAN = m_IBAN
End Property


Public Property Get StoricoLiquidazioniPath() As String
 StoricoLiquidazioniPath = m_StoricoLiquidazioniPath
End Property
Public Property Get StoricoECPath() As String
 StoricoECPath = m_StoricoECPath
End Property
Public Property Get StoricoECAdempiPath() As String
 StoricoECAdempiPath = m_StoricoECAdempiPath
End Property
Public Property Get StoricoEC_UNEP() As String
 StoricoEC_UNEP = m_StoricoEC_UNEP
End Property

Public Property Get LocalTempPath() As String
 LocalTempPath = m_LocalTempPath
End Property
Public Property Get BackupPath() As String
 BackupPath = m_AutoBackupPath
End Property
Public Property Get HistoryPath() As String
 HistoryPath = m_StoricoPath
End Property
Public Property Get dbPath() As String
 dbPath = m_dbPath
End Property
Public Property Get dbFile() As String
 dbFile = m_Dbname
End Property
Public Property Get CurrentDbFile() As String
 CurrentDbFile = m_CurrentDbName
End Property

Public Property Get ReportPath() As String
 ReportPath = m_reportPath
End Property


Private Sub LeggeRegistroPerDB(NomeFile As String)
        m_Dbname = Trim(GetSetting(NomeFile, "path", "database", "niente"))
        If m_Dbname = "niente" Then
            m_Dbname = "D\Atap\Atap.mdb"
            'chiamo la routine che scrive nel registro il percorso di default del mio DB
            ScriviRegistro NomeFile, "path", "database", m_Dbname
        End If

End Sub

Private Function LeggiRegistro(nomeapp As String, nomesez As String, nometag As String) As String
    LeggiRegistro = GetSetting(nomeapp, nomesez, nometag)
End Function

Private Sub ScriviRegistro(nomeapp As String, nomesez As String, nometag As String, valore As String)
    SaveSetting nomeapp, nomesez, nometag, valore
End Sub

Private Sub creaCartelle()
On Error Resume Next

m_LocalTempPath = app.Path & "\Temp"
If Dir(m_LocalTempPath, vbDirectory) = "" Then MkDir m_LocalTempPath

m_AutoBackupPath = m_dbPath & "AutoBackup"
If Dir(m_AutoBackupPath, vbDirectory) = "" Then MkDir m_AutoBackupPath

m_StoricoPath = m_dbPath & "Storico"
If Dir(m_StoricoPath, vbDirectory) = "" Then MkDir m_StoricoPath

m_StoricoLiquidazioniPath = m_dbPath & "Storico\Liquidazioni"
If Dir(m_StoricoLiquidazioniPath, vbDirectory) = "" Then MkDir m_StoricoLiquidazioniPath

m_StoricoECPath = m_dbPath & "Storico\EstrattiConto"
If Dir(m_StoricoECPath, vbDirectory) = "" Then MkDir m_StoricoECPath

m_StoricoECAdempiPath = m_dbPath & "Storico\EstrattiConto\Adempimenti"
If Dir(m_StoricoECAdempiPath, vbDirectory) = "" Then MkDir m_StoricoECAdempiPath

m_StoricoEC_UNEP = m_dbPath & "Storico\EstrattiConto\UNEP"
If Dir(m_StoricoEC_UNEP, vbDirectory) = "" Then MkDir m_StoricoEC_UNEP

On Error GoTo 0
err.Clear
End Sub
Private Function GetUserLock() As String
  GetUserLock = Left(ComputerName, 15) & Left(UserName, 15) & m_TIMESTAMP
End Function
Private Function ConnettiSQL() As ADODB.Connection
On Error GoTo DbErr

    Set ConnettiSQL = New ADODB.Connection
    ConnettiSQL.CursorLocation = adUseClient
    ConnettiSQL.Mode = adModeShareDenyNone
    ConnettiSQL.IsolationLevel = adXactReadCommitted
    ConnettiSQL.Provider = "sqloledb"

    ' Specify connection string on Open method.

    ConnettiSQL.Open "Server=PC1\SQLEXPRESS;Database=Atap;Trusted_Connection=yes"
      Exit Function
DbErr:
    Set ConnettiSQL = Nothing
    
    MsgBox " Err:" & err.Description & vbCrLf & dbFile, vbOKOnly + vbCritical, "Attenzione"
End Function


Private Function Connetti(dbFile As String) As ADODB.Connection
On Error GoTo DbErr
    
    Set Connetti = New ADODB.Connection
    Connetti.CursorLocation = adUseClient
    Connetti.Mode = adModeShareDenyNone
    Connetti.IsolationLevel = adXactReadCommitted
    m_CurrentDbName = dbFile
    If isFistAcces(dbFile) Then
     'DISABLE BACKUP
     ' Screen.MousePointer = vbHourglass
     ' BackUp m_AutoBackupPath
     ' Screen.MousePointer = vbDefault
    End If
    
    Connetti.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFile & ";;Jet OLEDB:Database;"
    
    Exit Function
DbErr:
    Set Connetti = Nothing
    
    MsgBox " Err:" & err.Description & vbCrLf & dbFile, vbOKOnly + vbCritical, "Attenzione"
    
    


End Function
Private Function isFistAcces(MDBFile As String)
 Dim fileName As String
 fileName = getPath(MDBFile) & "\" & GetFileNameWithoutExtension(MDBFile) & ".ldb"
 isFistAcces = (Dir(fileName) = "")
End Function

Private Sub BackUp(Path As String)
Dim sEXE As String, sPath1 As String, sPath2 As String, sDir As String, r As Integer
Dim fName As String
Dim tempFile As String
Dim fileName As String
On Error GoTo errback

 Screen.MousePointer = vbHourglass
 fileName = GetFileNameWithoutExtension(m_Dbname) & "_" & Weekday(Now)
 SafeKill m_LocalTempPath & "\TMP_*.mdb"
 tempFile = m_LocalTempPath & "\TMP_" & m_TIMESTAMP & ".mdb"
 Call CompactDatabase(m_Dbname, tempFile)
 
 If Dir(tempFile) <> "" Then
 
   ZipFile Path, fileName, m_LocalTempPath, "\TMP_" & m_TIMESTAMP & ".mdb"
 
 End If
 Screen.MousePointer = vbDefault
 
 Exit Sub
 
errback:
 MsgBox err.Description, vbOKOnly + vbExclamation
End Sub
