VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportFatturaRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public N_mov As String
Public N_riga As String
Public Causale As String
Public Valenza As String
Public Fiscale As String
Public DataRegistrazione As String
Public DataCompetenzaBil As String
Public DataCompetenzaIVA As String
Public DataDocumento As String
Public NumDocumento As String
Public DataDocRiferimento As String
Public DescrizioneMov As String
Public CodProvvisorio As String
Public DocIVAEsigibilDiff As String
Public RegimeIVAEsigibilDiff As String
Public TipoPagamento As String
Public ModPagamento As String
Public IncassoAccreditoDifferito As String
Public codPartita As String
Public Scadenze As String
Public NuovaScadenza As String
Public RegistroIVA As String
Public ProtocolloIVA As String
Public ProtocolloArrivo As String
Public RegIVAVendite As String
Public ImportoRitenuta As String
Public ImportoEnasarco As String
Public CodiceBarre As String
Public Allegato As String
Public CodProvenienza As String
Public CodiceFiscaleFunebri As String
Public CognomeNomeFunebri As String

Public TipoRiga As String
Public CodConto As String
Public CodNominativo As String
Public CFnominativo As String
Public DescrizioneRiga As String
Public ImportoRiga As String
Public ImportoIVA As String
Public ImportoValuta As String
Public CodValutaEstera As String
Public Cambio As String
Public codAliquotaIVA As String
Public codArticoloIVA As String
Public CodCentroCosto As String
Public CodAttISTAT As String
Public NumUnitProd As String
Public TipologiaSpesaSanitaria As String
Public TipoRettifica As String
Public ContoRettifica As String
Public DataCompBilancioRett As String
Public PeriodoRettDa As String
Public PeriodoRettA As String

Public Function CreateRow() As String

Dim s As String

    s = s & N_mov & ";"
    s = s & N_riga & ";"
    s = s & Causale & ";"
    s = s & Valenza & ";"
    s = s & Fiscale & ";"
    s = s & DataRegistrazione & ";"
    s = s & DataCompetenzaBil & ";"
    s = s & DataCompetenzaIVA & ";"
    s = s & DataDocumento & ";"
    s = s & NumDocumento & ";"
    s = s & DataDocRiferimento & ";"
    s = s & DescrizioneMov & ";"
    s = s & CodProvvisorio & ";"
    s = s & DocIVAEsigibilDiff & ";"
    s = s & RegimeIVAEsigibilDiff & ";"
    s = s & TipoPagamento & ";"
    s = s & ModPagamento & ";"
    s = s & IncassoAccreditoDifferito & ";"
    s = s & codPartita & ";"
    s = s & Scadenze & ";"
    s = s & NuovaScadenza & ";"
    s = s & RegistroIVA & ";"
    s = s & ProtocolloIVA & ";"
    s = s & ProtocolloArrivo & ";"
    s = s & RegIVAVendite & ";"
    s = s & ImportoRitenuta & ";"
    s = s & ImportoEnasarco & ";"
    s = s & CodiceBarre & ";"
    s = s & Allegato & ";"
    s = s & CodProvenienza & ";"
    s = s & CodiceFiscaleFunebri & ";"
    s = s & CognomeNomeFunebri & ";"
    
    s = s & TipoRiga & ";"
    s = s & CodConto & ";"
    s = s & CodNominativo & ";"
    s = s & CFnominativo & ";"
    s = s & DescrizioneRiga & ";"
    s = s & ImportoRiga & ";"
    s = s & ImportoIVA & ";"
    s = s & ImportoValuta & ";"
    s = s & CodValutaEstera & ";"
    s = s & Cambio & ";"
    s = s & codAliquotaIVA & ";"
    s = s & codArticoloIVA & ";"
    s = s & CodCentroCosto & ";"
    s = s & CodAttISTAT & ";"
    s = s & NumUnitProd & ";"
    s = s & TipologiaSpesaSanitaria & ";"
    s = s & TipoRettifica & ";"
    s = s & ContoRettifica & ";"
    s = s & DataCompBilancioRett & ";"
    s = s & PeriodoRettDa & ";"
    s = s & PeriodoRettA & ";"

    CreateRow = s
End Function
Public Sub Init(IsTestata As Boolean, numMovimento As String, numRiga As Long, dataFattura As String, numFattura As String, CF As String, PIVA As String, dImp As Double, dImpIVA As Double, sCodAliquota As String, sCodArticolo As String, desc As String, codCE As String)

 
  N_mov = numMovimento
  N_riga = numRiga
  DataRegistrazione = dataFattura
  ImportoRiga = Format(dImp, "##0.00")
  If dImpIVA = 0 Then
    ImportoIVA = ""
  Else
    ImportoIVA = Format(dImpIVA, "##0.00")
  End If
 If IsTestata Then
   DataDocumento = dataFattura
   NumDocumento = numFattura
   DescrizioneMov = "Fattura emessa"
   RegistroIVA = "V1"
   RegIVAVendite = "" ' "V1"
   ProtocolloIVA = "" 'numFattura
   Valenza = "0"
   Fiscale = "0"
   codPartita = numFattura & "|" & year(dataFattura)
   TipoRiga = "0"
   CodConto = g_Settings.CodTestata
   CFnominativo = IIf(PIVA <> "", PIVA, CF)
   ImportoIVA = ""
   Scadenze = dataFattura & "|" & ImportoRiga
 Else
   TipoRiga = "2"
   CodConto = codCE
   codAliquotaIVA = sCodAliquota
   codArticoloIVA = sCodArticolo
   DescrizioneMov = desc
 End If
End Sub
Private Sub Class_Initialize()
 Causale = "V05"
 
End Sub
