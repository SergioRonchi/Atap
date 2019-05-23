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
Private m_ExportPath As String
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
Private m_CodIVA As String

Private m_CodIVAEsente As String
Private m_CodArtIVAEsente As String
Private m_CodArtIVA As String

Private m_AtapUserFolder As String
Private m_AtapUserFolderTemp As String
Private m_AtapUserFolderBackup As String
Private m_Sede As String

Private m_LimiteBollo As Double
Private m_ImportoBollo As Double


Private m_CodTestata As String
Private m_CodTestataIncasso As String
Private m_CodQVar As String
Private m_CodQFissa As String
Private m_CodCompetenze As String

Private m_CurrentDbName As String
Private Const K_NomeInRegistry As String = "AtapDB"

Private m_StrutturaComuni As Collection



Private Sub Class_Initialize()
    
     Dim frm As frmPathDB
     Dim dbUpdater As CDBUpdater
     Set dbUpdater = New CDBUpdater
     m_TIMESTAMP = Format(Now, "YYMMDDhhmmss")
     m_USER_LOCK = GetUserLock

     m_ExportPath = GetSetting(K_NomeInRegistry, "Export", "Path", "")
     
     If Dir(m_ExportPath, vbDirectory) = "" Then
      m_ExportPath = ""
     End If
     
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
    
    Dim rs As ADODB.Recordset
    
    Set rs = newAdoRs
    
    rs.Open "SELECT * FROM Parametri", Me.DBConnection
    
    If Not rs.EOF Then
      m_IVA = rs("IVA")
      m_QuotaSoci = rs("QuotaSoci")
      m_LimiteSaldo = rs("LimiteSaldo")
      m_Banca = rs("Banca")
      m_IBAN = rs("IBAN")
      m_CodTestata = rs("CECodTestata")
       m_CodTestataIncasso = rs("CECodTestataIncasso")
      m_CodQVar = rs("CECodQVar")
      m_CodQFissa = rs("CECodQFixed")
      m_CodCompetenze = rs("CECodCompetenze")
      m_CodIVA = rs("CodIVA")
      
      m_CodIVAEsente = rs("CodIVAEsente")
      m_CodArtIVAEsente = rs("CodArtIVAEsente")
      m_CodArtIVA = rs("CodArtIVA")
      

      
      m_Sede = rs("Sede")
      m_LimiteBollo = rs("LimiteBollo")
      m_ImportoBollo = rs("ImportoBollo")

    End If
    
    
    
     
    m_AtapUserFolder = GetCommonUserFolder & "\Atap"
    SafeMakeDir m_AtapUserFolder
    m_AtapUserFolderTemp = m_AtapUserFolder & "\Temp"
    SafeMakeDir m_AtapUserFolderTemp
    m_AtapUserFolderBackup = m_AtapUserFolder & "\Backup"
    SafeMakeDir m_AtapUserFolderBackup
    
    Set m_StrutturaComuni = New Collection
 LoadComuniCollection app.Path & "\Dati\ListaComuni.txt", m_StrutturaComuni

End Sub
Public Property Get ExportPath() As String

 ExportPath = m_ExportPath
End Property
Public Property Let ExportPath(s As String)

 m_ExportPath = s
 SaveSetting K_NomeInRegistry, "Export", "Path", s
End Property
Public Sub ConnettiDB(sFile As String)
On Error Resume Next
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
Public Property Get iva() As Double
 iva = m_IVA
End Property
Public Property Let iva(x As Double)
  m_IVA = x
  Me.DBConnection.Execute "UPDATE Parametri SET IVA=" & Str(m_IVA)
End Property
Public Property Get QuotaSoci() As Double
 QuotaSoci = m_QuotaSoci
End Property
Public Property Let QuotaSoci(x As Double)
  m_QuotaSoci = x
  Me.DBConnection.Execute "UPDATE Parametri SET QuotaSoci=" & Str(m_QuotaSoci)
End Property
Public Property Let LimiteSaldo(x As Double)
  m_LimiteSaldo = x
  Me.DBConnection.Execute "UPDATE Parametri SET LimiteSaldo=" & Str(m_LimiteSaldo)
End Property

Public Property Let ImportoBollo(x As Double)
  m_ImportoBollo = x
  Me.DBConnection.Execute "UPDATE Parametri SET ImportoBollo=" & Str(m_ImportoBollo)
End Property

Public Property Let LimiteBollo(x As Double)
  m_LimiteBollo = x
  Me.DBConnection.Execute "UPDATE Parametri SET LimiteBollo=" & Str(m_LimiteBollo)
End Property




Public Property Get LimiteSaldo() As Double
 LimiteSaldo = m_LimiteSaldo
End Property


Public Property Get ImportoBollo() As Double
 ImportoBollo = m_ImportoBollo
End Property

Public Property Get LimiteBollo() As Double
 LimiteBollo = m_LimiteBollo
End Property



Public Property Let Banca(x As String)
  m_Banca = x
  Me.DBConnection.Execute "UPDATE Parametri SET Banca='" & Replace(m_Banca, "'", "''") & "'"
End Property
Public Property Get Banca() As String
 Banca = m_Banca
End Property
Public Property Let IBAN(x As String)
  m_IBAN = x
  Me.DBConnection.Execute "UPDATE Parametri SET IBAN='" & Replace(m_IBAN, "'", "''") & "'"
End Property
Public Property Get IBAN() As String
 IBAN = m_IBAN
End Property
Public Property Let CodIVA(x As String)
  m_CodIVA = x
  Me.DBConnection.Execute "UPDATE Parametri SET CodIVA='" & Replace(m_CodIVA, "'", "''") & "'"
End Property


Public Property Let CodIVAEsente(x As String)
  m_CodIVAEsente = x
  Me.DBConnection.Execute "UPDATE Parametri SET CodIVAEsente='" & Replace(m_CodIVAEsente, "'", "''") & "'"
End Property

Public Property Let CodArtIVA(x As String)
  m_CodArtIVA = x
  Me.DBConnection.Execute "UPDATE Parametri SET CodArtIVA='" & Replace(m_CodArtIVA, "'", "''") & "'"
End Property

Public Property Let CodArtIVAEsente(x As String)
  m_CodArtIVAEsente = x
  Me.DBConnection.Execute "UPDATE Parametri SET CodArtIVAEsente='" & Replace(m_CodArtIVAEsente, "'", "''") & "'"
End Property

Public Property Get CodIVA() As String
 CodIVA = m_CodIVA
End Property

Public Property Get CodIVAEsente() As String
 CodIVAEsente = m_CodIVAEsente
End Property

Public Property Get CodArtIVAEsente() As String
 CodArtIVAEsente = m_CodArtIVAEsente
End Property


Public Property Get CodArtIVA() As String
 CodArtIVA = m_CodArtIVA
End Property


Public Property Let Sede(x As String)
  m_Sede = x
  Me.DBConnection.Execute "UPDATE Parametri SET Sede='" & Replace(m_Sede, "'", "''") & "'"
End Property
Public Property Get Sede() As String
 Sede = m_Sede
End Property

Public Property Let CodTestata(x As String)
  m_CodTestata = x
  Me.DBConnection.Execute "UPDATE Parametri SET CECodTestata='" & Replace(m_CodTestata, "'", "''") & "'"
End Property
Public Property Get CodTestata() As String
 CodTestata = m_CodTestata
End Property
Public Property Get CodTestataIncasso() As String
 CodTestataIncasso = m_CodTestataIncasso
End Property
Public Property Let CodTestataIncasso(x As String)
  m_CodTestataIncasso = x
  Me.DBConnection.Execute "UPDATE Parametri SET CECodTestataIncasso='" & Replace(m_CodTestataIncasso, "'", "''") & "'"
End Property
Public Property Let CodQuataFissa(x As String)
  m_CodQFissa = x
  Me.DBConnection.Execute "UPDATE Parametri SET CECodQFixed='" & Replace(m_CodQFissa, "'", "''") & "'"
End Property
Public Property Get CodQuataFissa() As String
 CodQuataFissa = m_CodQFissa
End Property
Public Property Let CodQuotaVariabile(x As String)
  m_CodQVar = x
  Me.DBConnection.Execute "UPDATE Parametri SET CECodQVar='" & Replace(m_CodQVar, "'", "''") & "'"
End Property
Public Property Get CodQuotaVariabile() As String
 CodQuotaVariabile = m_CodQVar
End Property
Public Property Let CodCompetenze(x As String)
  m_CodCompetenze = x
  Me.DBConnection.Execute "UPDATE Parametri SET CECodCompetenze='" & Replace(m_CodCompetenze, "'", "''") & "'"
End Property
Public Property Get CodCompetenze() As String
 CodCompetenze = m_CodCompetenze
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
Public Property Get AtapUserFolder()
 AtapUserFolder = m_AtapUserFolder
 
End Property
Public Property Get AtapUserTempFolder()
 AtapUserTempFolder = m_AtapUserFolderTemp
 
End Property
Public Property Get AtapUserBackupFolder()
 AtapUserBackupFolder = m_AtapUserFolderBackup
 
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




Private Sub LoadComuniCollection(file As String, list As Collection)
     Dim Num
     Dim Fields
     Dim count As Long
     Dim riga As String
     Dim row As ComuneRecord
     Num = FreeFile
             
    Open file For Input As #Num
    
 
    While Not EOF(Num)
      Line Input #Num, riga
      If count > 0 Then
        Fields = Split(riga, ";")
        Set row = New ComuneRecord
        row.Instat = Fields(0)
        row.FieldName = Fields(1)
        row.Province = Fields(2)
        row.Region = Fields(3)
        row.Prefix = Fields(4)
        row.CAP = Fields(5)
        row.FiscalCode = Fields(6)
        row.People = Fields(7)
        row.Link = Fields(8)
        list.Add row, row.FiscalCode
        
      End If
      count = count + 1
    Wend
    
    Close Num
End Sub
Public Function GetComune(Code As String) As ComuneRecord
Dim o As ComuneRecord
On Error GoTo fine
Set o = m_StrutturaComuni(Code)
Set GetComune = o
fine:
End Function
