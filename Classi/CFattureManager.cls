VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFattureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetNextNumber() As Long
  GetNextNumber = GetADOValue("NumeriFattura", "Numero", "1=1", g_Settings.DBConnection, True) + 1
End Function

Public Function SaveNumber(value As Long) As Long
  g_Settings.DBConnection.Execute "UPDATE NumeriFattura SET Numero=" & value
End Function


