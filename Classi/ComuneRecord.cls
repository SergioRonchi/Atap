VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComuneRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Instat As String
Private m_Name As String
Private m_Provincia As String
Private m_Regione As String
Private m_Prefix As String

Private m_CAP As String
Private m_FiscalCode As String
Private m_People As Long
Private m_Link As String


Public Property Get Instat() As String
   Instat = m_Instat
End Property
Public Property Let Instat(v As String)
   m_Instat = v
End Property

Public Property Get FieldName() As String
   FieldName = m_Name
End Property
Public Property Let FieldName(v As String)
   m_Name = v
End Property

Public Property Get Province() As String
   Province = m_Provincia
End Property
Public Property Let Province(v As String)
   m_Provincia = v
End Property

Public Property Get Region() As String
   Region = m_Regione
End Property
Public Property Let Region(v As String)
   m_Regione = v
End Property

Public Property Get Prefix() As String
   Prefix = m_Prefix
End Property
Public Property Let Prefix(v As String)
   m_Prefix = v
End Property

Public Property Get CAP() As String
   CAP = m_CAP
End Property
Public Property Let CAP(v As String)
   m_CAP = v
End Property

Public Property Get FiscalCode() As String
   FiscalCode = m_FiscalCode
End Property
Public Property Let FiscalCode(v As String)
   m_FiscalCode = v
End Property
Public Property Get People() As Long
   People = m_People
End Property
Public Property Let People(v As Long)
   m_People = v
End Property
Public Property Get Link() As String
   Link = m_Link
End Property
Public Property Let Link(v As String)
   m_Link = v
End Property

