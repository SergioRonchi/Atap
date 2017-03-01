Attribute VB_Name = "Procedure"
Option Explicit

Public Sub CaricaDayNav(cmb As Object)
 cmb.AddItem "Oggi"
 cmb.AddItem "DaIeri"
 cmb.AddItem "QuestaSettimana"
 cmb.AddItem "DaUltimaSettimana"
 cmb.AddItem "QuestoMese"
 cmb.AddItem "UltimoMese"
 cmb.AddItem "AnnoCorrente"
 cmb.AddItem "DaUltimoAnno"
 cmb.AddItem "InteroAnno"
 cmb.AddItem "InteroAnnoPrec"
 
End Sub
Public Property Get GLO_OGGI_OLD() As String
   
   GLO_OGGI_OLD = Format(Date, "yyyy") - 1 & Right(GLO_OGGI, 4)
End Property
Public Property Get GLO_OGGI() As String
   GLO_OGGI = Format(Date, "yyyymmdd")
End Property
Public Sub SettaDate(Dal As Object, Al As Object, Scelta As Integer)
Dim d As Date, x As Integer
Dim Mese As String, Giorno As String, gset As Integer
Dim MesePrec As String

Dim m_SysDataFormat As String
Dim SysShortData As String
Dim m_SysLang As String
Dim m_SysCountry As String
Dim m_SysValuta As String
Dim m_SysDecimale As String
Dim m_SysSeparatore As String
Dim m_RegionalData As String
d = Date

x = month(d)
MesePrec = x - 1
If MesePrec < 1 Then MesePrec = 12

Mese = IIf(x < 10, "0" & x, x)

x = day(d)
Giorno = IIf(x < 10, "0" & x, x)

gset = Weekday(d, vbMonday)

 Call GetTheLocaleInfo(m_SysDataFormat, SysShortData, m_SysLang, m_SysCountry, m_SysValuta, m_SysDecimale, m_SysSeparatore)

Select Case Scelta
  Case 0 'Oggi
       Dal = getRegionalData(GLO_OGGI, SysShortData, True)
       Al = getRegionalData(GLO_OGGI, SysShortData, True)
  Case 1 'DA ieri
       Dal = getRegionalData(Format(d - 1, "yyyymmdd"), SysShortData, True)
       Al = getRegionalData(GLO_OGGI, SysShortData, True)
  Case 2 'Questa settimana
        Dal = getRegionalData(Format(d - gset + 1, "yyyymmdd"), SysShortData, True)
        Al = getRegionalData(GLO_OGGI, SysShortData, True)
  Case 3 'Dall ' ultima settimana
        Dal = getRegionalData(Format(d - 7 - gset + 1, "yyyymmdd"), SysShortData, True)
        Al = getRegionalData(GLO_OGGI, SysShortData, True)
  Case 4 'Questo Mese
        Dal = getRegionalData(Format(d - Giorno + 1, "yyyymmdd"), SysShortData, True)
        
        Al = getRegionalData(Format(LastDayOfMonth(month(Date), year(Date)), "YYYYMMDD"), SysShortData, True)
  Case 5 'Dall ' ultimo mese
        Dal = getRegionalData(Format("01/" & MesePrec & "/" & year(d), "yyyymmdd"), SysShortData, True)
        Al = getRegionalData(GLO_OGGI, SysShortData, True)
  
  Case 6 'Anno corrente  (si intende da inizio anno a oggi)
        Dal = getRegionalData(year(d) & "0101", SysShortData, True)
        Al = getRegionalData(GLO_OGGI, SysShortData, True)
  
  Case 7 'Dall ' anno prededente
        Dal = getRegionalData(GLO_OGGI_OLD, SysShortData, True)
        Al = getRegionalData(GLO_OGGI, SysShortData, True)
  Case 8 'Intero Anno
        Dal = getRegionalData(year(d) & "0101", SysShortData, True)
        Al = getRegionalData(year(d) & "1231", SysShortData, True)
  Case 9 'Intero Anno Prec
        Dal = getRegionalData(year(d) - 1 & "0101", SysShortData, True)
        Al = getRegionalData(year(d) - 1 & "1231", SysShortData, True)
        
  Case 10
        'Prossimi sette giorni
        Dal = getRegionalData(GLO_OGGI, SysShortData, True)
        Al = getRegionalData(Format(d + 7, "yyyymmdd"), SysShortData, True)
        
  Case 11
        'prossimi 30 giorni
        Dal = getRegionalData(GLO_OGGI, SysShortData, True)
        Al = getRegionalData(Format(d + 30, "yyyymmdd"), SysShortData, True)
        

     

End Select

     Al.ForeColor = &H80000008
     Dal.ForeColor = &H80000008
End Sub
Public Function getRegionalData(field As String, formato As String, Optional isValue As Boolean) As String
If Not isValue Then
   getRegionalData = " Format( Mid(" & field & ",1,4) & '-' & Mid(" & field & ",5,2) & '-' & Mid(" & field & ",7,2),'" & formato & "') "
'   Select Case formato
'    Case "mm/dd/yyyy"
'     getRegionalData = " Mid(" & field & ",5,2) & '/' & Mid(" & field & ",7,2) & '/' & Mid(" & field & ",1,4) "
'    Case "dd/mm/yyyy"
'     getRegionalData = " Mid(" & field & ",7,2) & '/' & Mid(" & field & ",5,2) & '/' & Mid(" & field & ",1,4) "
'   End Select
  Else
   Dim d As Date
   d = Mid(field, 1, 4) & "-" & Mid(field, 5, 2) & "-" & Mid(field, 7, 2)
   getRegionalData = Format(d, formato)
End If
     
End Function
Public Sub sortGrid(flex As VSFlexGrid, Col As Long, ByRef Order As Integer, oldCol As Integer, newCol As Integer)
    
   
    If newCol < 0 Then newCol = flex.ColIndex("NumOrdinamento")
    ' no flags? apply custom sort
    If flex.ExplorerBar > &H1000& Then Exit Sub
    
    ' save selection
    Dim r&, c&, rs&, cs&
    flex.GetSelection r, c, rs, cs
    flex.Redraw = flexRDNone
    
    ' apply sort to non-empty range
    Dim row%
    For row = flex.Rows - 1 To flex.FixedRows Step -1
        If Len(flex.TextMatrix(row, Col)) Then Exit For
    Next
    If row > flex.FixedRows Then
        flex.Select flex.FixedRows, Col, row, Col
        flex.Sort = Order
    End If
    
    If Col = oldCol Then
        flex.Select flex.FixedRows, newCol, flex.Rows - 1, newCol
        flex.Sort = Order

    End If
    
    ' restore selection
    flex.Select r, c, rs, cs
    flex.Redraw = flexRDDirect
    
    ' cancel default sort
    Order = 0

End Sub
Public Function ControlloNULL(DATO As Variant) As String

On Error GoTo ControlloNullError

If IsNull(DATO) Then
    ControlloNULL = " "
Else
    ControlloNULL = DATO
End If

Exit Function

ControlloNullError:
    
    Resume Next

End Function


Public Function controlloPIvaCodFis(cod As String) As Boolean
    
    Dim err As Boolean
    Dim d As String
    Dim n As Integer
    Dim k As Integer
    Dim x As Integer
    Dim I As Integer
    n = k = x = 0
    Dim Y(43) As Integer
          
    err = False
    
    Y(1) = 1
    Y(2) = 0
    Y(3) = 5
    Y(4) = 7
    Y(5) = 9
    For I = 6 To 10
        Y(I) = (2 * I + 1)
    Next I
    For I = 11 To 17
        Y(I) = 0
    Next I
    For I = 18 To 27
        Y(I) = Y(I - 17)
    Next I
    Y(28) = 2
    Y(29) = 4
    Y(30) = 18
    Y(31) = 20
    Y(32) = 11
    Y(33) = 3
    Y(34) = 6
    Y(35) = 8
    Y(36) = 12
    Y(37) = 14
    Y(38) = 16
    Y(39) = 10
    Y(40) = 22
    Y(41) = 25
    Y(42) = 24
    Y(43) = 23

    If Len(cod) = 11 Then
        n = 0
        err = True
        For I = 1 To 10 Step 2
          n = n + Asc(Mid(cod, I, 1))
        Next I
        For I = 2 To 10 Step 2
            k = Asc(Mid(cod, I, 1))
            If k > 9 Then
                k = k / 10 + k Mod 10
            End If
            n = n + k
        Next I
        x = (10 - (n Mod 10)) Mod 10
        If Asc(Mid(cod, 11, 1)) Then
            err = False
        End If
    Else
        err = True
        For I = 2 To 15 Step 2
            x = Asc(Mid(cod, I, 1))
            If x > 60 Then
                n = n + x - 65
            Else
                n = n + x - 48
            End If
        Next I
        
        For I = 1 To 15 Step 2
            x = Asc(Mid(cod, I, 1)) - 47
            n = n + Y(x)
        Next I
        n = n Mod 26
        d = Chr(n + 65)
        If Mid(cod, I, 1) = d Then
            err = False
        End If
    End If
        
    controlloPIvaCodFis = err

End Function
Public Function valore(v As String) As Double
 v = Format(v, "#.##")
 valore = val(Replace(v, ",", "."))
 
End Function
Public Function FixDouble(v As Double) As String
 Dim s As String
 s = CStr(v)

 FixDouble = Replace(s, ",", ".")
 
End Function
Public Sub PopolaTDBCombo(c As TDBCombo, Tabella As String, field As String, Optional codice As String, Optional Tutti As Boolean, Optional showCode As Boolean, Optional OrderBy As String = "", Optional SQL As String = "")
Dim rs As ADODB.Recordset
Dim campi As String
Dim I As Integer


campi = field
If codice <> "" Then campi = campi & "," & codice
If SQL = "" Then
    SQL = "SELECT " & campi & " FROM " & Tabella
    If OrderBy <> "" Then
      SQL = SQL & " ORDER BY " & OrderBy
    End If
End If
If Tutti Then
 SQL = "SELECT First('- Mostra Tutto -') as [" & field & "],'XXALLXX' FROM " & Tabella & " UNION ALL " & SQL
End If
Set rs = newAdoRs
rs.Open SQL, g_Settings.DBConnection
c.Clear
Set c.RowSource = rs
If Tutti Then
  
    c.Text = "- Mostra Tutto -"
    
End If


  If Not showCode Then c.Columns(1).Visible = False
  

End Sub
Public Sub PopolaCombo(c As ComboBox, Tabella As String, field As String, Optional codice As String, Optional ByRef Col As Collection, Optional Tutti As Boolean)
'Ronchi 12 giugno 2002
On Error GoTo fine
Dim rs As ADODB.Recordset
Set rs = newAdoRs
rs.Open Tabella, g_Settings.DBConnection
c.Clear
Set Col = New Collection
 If Tutti Then
   c.AddItem "- Mostra Tutto -"
   Col.Add "TUTTO"
 End If
While Not rs.EOF
  c.AddItem rs(field)
  If codice <> "" Then Col.Add rs(codice).value
  rs.MoveNext
Wend
c.ListIndex = 0
Exit Sub
fine:
On Error GoTo 0

End Sub


Public Function RitornaData(Data As String) As String

Dim anno, Mese, Giorno As String

If Data = " " Or Data = "" Then
    RitornaData = Data
Else
    anno = Mid(Data, 1, 4)
    Mese = Mid(Data, 5, 2)
    Giorno = Mid(Data, 7, 2)
    RitornaData = Giorno + "/" + Mese + "/" + anno
End If

End Function

Public Function ElaboraData(Data As String) As String

Dim anno, Mese, Giorno As String

If Data = " " Or Data = "" Then
    ElaboraData = Data
Else
    If InStr(1, Data, "/") Then Data = Replace(Data, "/", "")
    Giorno = Mid(Data, 1, 2)
    Mese = Mid(Data, 3, 2)
    anno = Mid(Data, 5, 4)
    ElaboraData = anno + Mese + Giorno
End If

End Function

Public Sub Ridimensiona(frm As Form, tipo As String)
 With frm
  .fraTop.Top = 0
  Select Case tipo
     Case "small"
           .Height = .fraTop.Top + .fraTop.Height + .fraComandi.Height + 400
           .fraComandi.Top = .fraTop.Height + .fraTop.Top
           .fraMain.Visible = False
     Case "big"
           .Height = .fraTop.Top + .fraTop.Height + .fraComandi.Height + .fraMain.Height + 400
           .fraMain.Top = .fraTop.Height + .fraTop.Top
           .fraComandi.Top = .fraMain.Height + .fraMain.Top
           .fraMain.Visible = True
  End Select
 End With
End Sub



Public Function Approssima(MiaValuta As Double) As Double
    Dim MioVal1 As Double
    Dim MioVal2 As Double
    Dim MyStr As String
    Dim MyStr1 As String
    Dim MyStr2 As String
    Dim MiaString As String
    Dim HaDecimali As Boolean
    Dim I As Integer
    HaDecimali = False
    If MiaValuta <> 0 Then
        MiaString = Trim(Str(MiaValuta))
        For I = 1 To Len(MiaString)
            MyStr = Mid(MiaString, I, 1)
            If MyStr = "." Then
                HaDecimali = True
                MyStr1 = Left(MiaString, I - 1)
                If (Len(Trim(MyStr1)) = 0) Then
                    MyStr1 = "0"
                End If
                MyStr2 = Right(MiaString, Len(MiaString) - I)
                I = I + 1
            End If
        Next I
        If HaDecimali = False Then
            Approssima = MiaString
            Exit Function
        End If
        MioVal1 = CDbl(MyStr1)
        MioVal2 = CDbl(Left(MyStr2, 1))
        If MioVal2 < 5 Then
            Approssima = MioVal1
        Else
            Approssima = MioVal1 + 1
        End If
    Else
        Approssima = 0
    End If
End Function
