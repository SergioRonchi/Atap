VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportAnagraficaRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public codice As String
Public CodiceFiscale As String
Public FLCliente As String
Public FLCAF As String
Public FLIntermediario As String
Public IDTipologia As String
Public COGNOME As String
Public nome As String
Public Denominazione As String
Public Sesso As String
Public DataNascita As String
Public IDStatoCivile As String
Public ComuneNascita As String
Public IDProvinciaNascita As String
Public CodComuneNascita As String
Public PartitaIVA As String
Public Sigla As String
Public CodiceNaturaGiuridica As String

Public IndirizzoResidenza As String
Public NumeroCivicoResidenza As String
Public ComuneResidenza As String
Public IDProvinciaResidenza As String
Public CodComuneResidenza As String
Public CAPResidenza As String
Public FLPrincipale As String

Public NazioneEstera As String
Public CodCatastNazione As String
Public CodEsteroNazione As String
Public StatoFederaleEstero As String
Public LocalitaEstero As String
Public IndirizzoEstero As String
Public NumeroCivicoEstero As String
Public ZIP As String
Public FLNazionalitaEstera As String
Public FLCittadinanzaItalia As String
Public CFEstero As String
Public FLPrincipaleEstero As String
Public PartitaIVAEstera As String
Public IDISO As String
Public CFStabile As String

Public IndirizzoDomFisc As String
Public NumeroCivico As String
Public ComuneDomFisc As String
Public IDProvinciaDomFisc As String
Public CodComuneDomFisc As String
Public CAPDomFisc As String
Public Domfiscale_indirizzo_primario As String

Public telefono As String
Public fax As String
Public cellulare As String
Public Indirizzo_email As String
Public Indirizzo_URL As String

Public Data_inizio_rapporto As String
Public Note As String
Public CodiceFiscaleAzienda As String
Public FLCliFor As String


Public Function CreateRow() As String

Dim s   As String

    s = s & codice & ";"
 
    s = s & CodiceFiscale & ";"
    s = s & FLCliente & ";"
    s = s & FLCAF & ";"
    s = s & FLIntermediario & ";"
    s = s & IDTipologia & ";"
    s = s & COGNOME & ";"
    s = s & nome & ";"
    s = s & Denominazione & ";"
    s = s & Sesso & ";"
    s = s & DataNascita & ";"
    s = s & IDStatoCivile & ";"
    s = s & ComuneNascita & ";"
    s = s & IDProvinciaNascita & ";"
    s = s & CodComuneNascita & ";"
    s = s & PartitaIVA & ";"
    s = s & Sigla & ";"
    s = s & CodiceNaturaGiuridica & ";"
    
    s = s & IndirizzoResidenza & ";"
    s = s & NumeroCivicoResidenza & ";"
    s = s & ComuneResidenza & ";"
    s = s & IDProvinciaResidenza & ";"
    s = s & CodComuneResidenza & ";"
    s = s & CAPResidenza & ";"
    s = s & FLPrincipale & ";"
    
    s = s & NazioneEstera & ";"
    s = s & CodCatastNazione & ";"
    s = s & CodEsteroNazione & ";"
    s = s & StatoFederaleEstero & ";"
    s = s & LocalitaEstero & ";"
    s = s & IndirizzoEstero & ";"
    s = s & NumeroCivicoEstero & ";"
    s = s & ZIP & ";"
    s = s & FLNazionalitaEstera & ";"
    s = s & FLCittadinanzaItalia & ";"
    s = s & CFEstero & ";"
    s = s & FLPrincipaleEstero & ";"
    s = s & PartitaIVAEstera & ";"
    s = s & IDISO & ";"
    s = s & CFStabile & ";"
    
    s = s & IndirizzoDomFisc & ";"
    s = s & NumeroCivico & ";"
    s = s & ComuneDomFisc & ";"
    s = s & IDProvinciaDomFisc & ";"
    s = s & CodComuneDomFisc & ";"
    s = s & CAPDomFisc & ";"
    s = s & Domfiscale_indirizzo_primario & ";"
    
    s = s & telefono & ";"
    s = s & fax & ";"
    s = s & cellulare & ";"
    s = s & Indirizzo_email & ";"
    s = s & Indirizzo_URL & ";"
    
    s = s & Data_inizio_rapporto & ";"
    s = s & Note & ";"
    s = s & CodiceFiscaleAzienda & ";"
    s = s & FLCliFor
    
    CreateRow = s

End Function
Public Sub Init(sDenominazione As String, CF As String, PIVA As String, _
                sIndirizzoResidenza As String, sLocalitaResidenza As String, sCAPResidenza As String, sCivicoResidenza As String, sProvResidenza As String, _
                sTelefono As String, sCellulare As String, sEmail As String, sFax As String, _
                sNota As String)
                
 Dim isAzienda As Boolean
 
 isAzienda = PIVA <> "" And IsNumeric(Right(CF, 1))
 IDTipologia = "9"
 Denominazione = sDenominazione
 CodiceFiscale = CF
 PartitaIVA = PIVA
 IndirizzoResidenza = sIndirizzoResidenza
 ComuneResidenza = sLocalitaResidenza
 CAPResidenza = sCAPResidenza
 IDProvinciaResidenza = sProvResidenza
 NumeroCivicoResidenza = sCivicoResidenza
 
 telefono = sTelefono
 cellulare = sCellulare
 Indirizzo_email = sEmail
 fax = sFax
 
 Note = sNota
 
 
 Dim SX As String
 Dim DN As String
 Dim PN As String
 Dim CN As String
 If CF <> "" And Not IsNumeric(Right(CF, 1)) Then
   GetFROMFiscalCode CF, SX, DN, PN, CN
   Sesso = SX
   DataNascita = DN
   IDProvinciaNascita = PN
   ComuneNascita = CN
   
 End If
   
 
End Sub


Private Sub GetFROMFiscalCode(cfisc As String, ByRef SX As String, ByRef DN As String, ByRef PN As String, ByRef CN As String)
On Error GoTo fine
 Dim year As String
 Dim month As String
 Dim day As String
 Dim Code As String
 
 year = Mid(cfisc, 7, 2)
 If year > 50 Then
   year = 19 & year
  Else
   year = 20 & year
 End If
 month = Mid(cfisc, 9, 1)
 Select Case Mid(cfisc, 9, 1)
   Case "A"
     month = "01"
   Case "B"
     month = "02"
   Case "C"
        month = "03"
   Case "D"
        month = "04"
   Case "E"
        month = "05"
   Case "H"
        month = "06"
   Case "L"
        month = "07"
   Case "M"
        month = "08"
   Case "P"
        month = "09"
   Case "R"
        month = "10"
   Case "S"
        month = "11"
   Case "T"
        month = "12"
 End Select
 day = Mid(cfisc, 10, 2)
 If day > 35 Then
    SX = "F"
    day = day - 40
    day = String(2 - Len(day), "0") & day
  Else
    SX = "M"
 End If
 DN = day & "/" & month & "/" & year
 Code = Mid(cfisc, 12, 4)
 Dim o As ComuneRecord
 Set o = g_Settings.GetComune(Code)
 If Not o Is Nothing Then
   CN = o.FieldName
   PN = o.Province
  Else
    CN = ""
    PN = ""
 End If
fine:
End Sub
