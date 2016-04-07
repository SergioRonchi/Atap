Attribute VB_Name = "Registro"
Option Explicit

Sub LeggeRegistroPerDB(NomeFile As String)
        gDbName = Trim(GetSetting(NomeFile, "path", "database", "niente"))
        If gDbName = "niente" Then
            gDbName = "D\Atap\Atap.mdb"
            'chiamo la routine che scrive nel registro il percorso di default del mio DB
            Call ScriviRegistro(NomeFile, "path", "database", gDbName)
        End If

End Sub

Private Function LeggiRegistro(nomeapp As String, nomesez As String, nometag As String) As String
    LeggiRegistro = GetSetting(nomeapp, nomesez, nometag)
End Function

Private Sub ScriviRegistro(nomeapp As String, nomesez As String, nometag As String, valore As String)
'scrivo nel registro il percorso del mio DB

'nomeapp contiene il nome dell'applicazione o del progetto;
'nomesez contiene la stringa con il nome della sezione nella quale viene salvata l'impostazione di chiave;
'nometag contiene stringa con il nome dell'impostazione di chiave salvata;
    SaveSetting nomeapp, nomesez, nometag, valore
End Sub


