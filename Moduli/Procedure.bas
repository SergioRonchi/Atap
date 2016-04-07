Attribute VB_Name = "Procedure"
Option Explicit
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
    Dim D As String
    Dim n As Integer
    Dim k As Integer
    Dim X As Integer
    Dim i As Integer
    n = k = X = 0
    Dim Y(43) As Integer
          
    err = False
    
    Y(1) = 1
    Y(2) = 0
    Y(3) = 5
    Y(4) = 7
    Y(5) = 9
    For i = 6 To 10
        Y(i) = (2 * i + 1)
    Next i
    For i = 11 To 17
        Y(i) = 0
    Next i
    For i = 18 To 27
        Y(i) = Y(i - 17)
    Next i
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
        For i = 1 To 10 Step 2
          n = n + Asc(Mid(cod, i, 1))
        Next i
        For i = 2 To 10 Step 2
            k = Asc(Mid(cod, i, 1))
            If k > 9 Then
                k = k / 10 + k Mod 10
            End If
            n = n + k
        Next i
        X = (10 - (n Mod 10)) Mod 10
        If Asc(Mid(cod, 11, 1)) Then
            err = False
        End If
    Else
        err = True
        For i = 2 To 15 Step 2
            X = Asc(Mid(cod, i, 1))
            If X > 60 Then
                n = n + X - 65
            Else
                n = n + X - 48
            End If
        Next i
        
        For i = 1 To 15 Step 2
            X = Asc(Mid(cod, i, 1)) - 47
            n = n + Y(X)
        Next i
        n = n Mod 26
        D = Chr(n + 65)
        If Mid(cod, i, 1) = D Then
            err = False
        End If
    End If
        
    controlloPIvaCodFis = err

End Function
Public Function valore(v As String) As Double
 v = Format(v, "#.##")
 valore = val(Replace(v, ",", "."))
 
End Function
Public Sub PopolaTDBCombo(c As TDBCombo, Tabella As String, Field As String, Optional codice As String, Optional Tutti As Boolean, Optional showCode As Boolean, Optional OrderBy As String = "", Optional SQL As String = "")
Dim rs As ADODB.Recordset
Dim campi As String
Dim i As Integer


campi = Field
If codice <> "" Then campi = campi & "," & codice
If SQL = "" Then
    SQL = "SELECT " & campi & " FROM " & Tabella
    If OrderBy <> "" Then
      SQL = SQL & " ORDER BY " & OrderBy
    End If
End If
If Tutti Then
 SQL = "SELECT First('- Mostra Tutto -') as [" & Field & "],'XXALLXX' FROM " & Tabella & " UNION ALL " & SQL
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
Public Sub PopolaCombo(c As ComboBox, Tabella As String, Field As String, Optional codice As String, Optional ByRef Col As Collection, Optional Tutti As Boolean)
'Ronchi 12 giugno 2002
On Error GoTo FINE
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
  c.AddItem rs(Field)
  If codice <> "" Then Col.Add rs(codice).value
  rs.MoveNext
Wend
c.ListIndex = 0
Exit Sub
FINE:
On Error GoTo 0

End Sub


Public Function RitornaData(data As String) As String

Dim anno, mese, giorno As String

If data = " " Or data = "" Then
    RitornaData = data
Else
    anno = Mid(data, 1, 4)
    mese = Mid(data, 5, 2)
    giorno = Mid(data, 7, 2)
    RitornaData = giorno + "/" + mese + "/" + anno
End If

End Function

Public Function ElaboraData(data As String) As String

Dim anno, mese, giorno As String

If data = " " Or data = "" Then
    ElaboraData = data
Else
    If InStr(1, data, "/") Then data = Replace(data, "/", "")
    giorno = Mid(data, 1, 2)
    mese = Mid(data, 3, 2)
    anno = Mid(data, 5, 4)
    ElaboraData = anno + mese + giorno
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
    Dim i As Integer
    HaDecimali = False
    If MiaValuta <> 0 Then
        MiaString = Trim(Str(MiaValuta))
        For i = 1 To Len(MiaString)
            MyStr = Mid(MiaString, i, 1)
            If MyStr = "." Then
                HaDecimali = True
                MyStr1 = Left(MiaString, i - 1)
                If (Len(Trim(MyStr1)) = 0) Then
                    MyStr1 = "0"
                End If
                MyStr2 = Right(MiaString, Len(MiaString) - i)
                i = i + 1
            End If
        Next i
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
