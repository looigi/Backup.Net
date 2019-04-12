Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Ionic.Zip

Public Class OperazioniSuFileDettagli
    Private FileOrigine As String
    Private FileDestinazione As String
    Private gf As GestioneFilesDirectory
    Private log As StringBuilder = New StringBuilder

    Sub New()
        gf = New GestioneFilesDirectory
    End Sub

    Protected Overrides Sub Finalize()
        gf = Nothing

        MyBase.Finalize()
    End Sub

    Public Function fAttendi(idProc As Integer, Origine As String) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "ATTENDI: " & Origine & " secondi")

        log = New StringBuilder

        Attendi(Origine)

        ope.ScriveLogServizio(idProc, "USCITA ATTENDI")
        ope = Nothing

        Return log
    End Function

    Public Function AvviaEseguibile(idProc As Integer, Origine As String, Parametro As String) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "AVVIA ESEGUIBILE: " & Origine)

        log = New StringBuilder

        'Dim parametro As String = Origine
        'If parametro.Contains(" ") Then
        '    parametro = Mid(parametro, parametro.IndexOf(" ") + 1, parametro.Length).Trim
        '    Origine = Origine.Replace(parametro, "").Trim
        'Else
        '    parametro = ""
        'End If

        Try
            If Parametro = "" Then
                Process.Start(Origine)
            Else
                Process.Start(Origine, Parametro)
            End If
        Catch ex As Exception
            ope.ScriveLogServizio(idProc, "AVVIA ESEGUIBILE: Errore " & ex.Message)
        End Try

        gf.CreaAggiornaFile("PassaggioService.Dat", "AVVIO APPLICAZIONE;" & Origine & ";")

        ope.ScriveLogServizio(idProc, "USCITA AVVIA ESEGUIBILE")
        ope = Nothing

        Return log
    End Function

    Public Function FermaEseguibile(idProc As Integer, Origine As String) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "STOPPA ESEGUIBILE: " & Origine)

        log = New StringBuilder

        FermaServizio(idProc, Origine)

        gf.CreaAggiornaFile("PassaggioService.Dat", "CHIUSURA APPLICAZIONE;" & Origine & ";")

        ope.ScriveLogServizio(idProc, "USCITA STOPPA ESEGUIBILE")
        ope = Nothing

        Return log
    End Function

    Public Function RiavvioPC(idProc As Integer) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "RIAVVIO PC:")

        log = New StringBuilder

        Try
            FaiPartireServizio(idProc, "BackupLauncher.NET", "")

            gf.CreaAggiornaFile("PassaggioService.Dat", "RIAVVIO;")

            ope.ScriveLogServizio(idProc, "USCITA RIAVVIO PC")
        Catch ex As Exception
            log.Append("ERRORE: " & ex.Message)
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function EliminaFile(idProc As Integer, Origine As String, Filtro As String, SottoDirectory As String, lblOperazione As Label, lblContatore As Label) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "ELIMINAZIONE FILE:")

        log = New StringBuilder

        Try
            If SottoDirectory = "S" Then
                gf.ScansionaDirectorySingola(Origine, Filtro, lblOperazione, False)
            Else
                gf.ScansionaDirectorySingola(Origine, Filtro, lblOperazione, True)
            End If

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    FileOrigine = Filetti(i)

                    gf.EliminaFileFisico(FileOrigine)
                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Elimina file " & gf.TornaNomeFileDaPath(Filetti(i)), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False))
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function EliminazioneDirectory(idProc As Integer, Origine As String,
                                          lblOperazione As Label, lblContatore As Label) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "ELIMINAZIONE DIRECTORY:")

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, "", lblOperazione, False)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati
            Dim Cartelle() As String = gf.RitornaDirectoryRilevate
            Dim qCartelle As Long = gf.RitornaQuanteDirectoryRilevate

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    'FileDestinazione = Origine & "\" & gf.TornaNomeFileDaPath(Filetti(i))

                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione file in directory " & gf.TornaNomeFileDaPath(Filetti(i)), i & "/" & qFiletti)

                    gf.EliminaFileFisico(Filetti(i))
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            If Not BloccaTutto Then
                For k As Integer = 1 To 3
                    For i As Long = qCartelle To 0 Step -1
                        If Cartelle(i) <> "" Then
                            Try
                                RmDir(Cartelle(i))
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione directory " & Cartelle(i), i & "/" & qCartelle)
                            Catch ex As Exception
                            End Try
                        End If

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If BloccaTutto Then
                            Exit For
                        End If
                    Next

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    If BloccaTutto Then
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function CreaDirectory(idProc As Integer, Origine As String, Filtro As String, lblOperazione As Label, lblContatore As Label) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "CREAZIONE DIRECTORY:")

        log = New StringBuilder

        Try
            FileOrigine = Origine & "\" & Filtro

            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Creazione directory " & FileOrigine, " ")

            gf.CreaDirectoryDaPercorso(FileOrigine & "\")
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function Copia(idProc As Integer, Operazione As Integer, Origine As String, Destinazione As String, qFiletti As Long,
                          Filetti() As String, Sovrascrivi As String, lblOperazione As Label, lblContatore As Label) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "SPOSTAMENTO FILES: " & Origine & "->" & Destinazione)

        log = New StringBuilder

        Try
            Dim PathUlteriore As String = ""

            gf.ScansionaDirectorySingola(Origine, "", lblOperazione, False)

            Filetti = gf.RitornaFilesRilevati
            qFiletti = gf.RitornaQuantiFilesRilevati

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    FileOrigine = Filetti(i)
                    PathUlteriore = FileOrigine.Replace(Origine & "\", "")
                    FileDestinazione = Destinazione & "\" & PathUlteriore ' & "\" & gf.TornaNomeFileDaPath(Filetti(i))

                    gf.CreaDirectoryDaPercorso(Destinazione & "\")
                    gf.CopiaFileFisico(FileOrigine, FileDestinazione, IIf(Sovrascrivi = "S", True, False))

                    If Operazione = TipoOperazione.Spostamento Then
                        gf.EliminaFileFisico(FileOrigine)
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Sposta file " & gf.TornaNomeFileDaPath(Filetti(i)), i & "/" & qFiletti)
                    Else
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Copia file " & gf.TornaNomeFileDaPath(Filetti(i)), i & "/" & qFiletti)
                    End If
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function Sincronizza(idProc As Integer, Progressivo As Integer, Db As GestioneACCESS, ConnSQL As Object, Origine As String, Destinazione As String, Filtro As String,
                                lblOperazione As Label, lblContatore As Label, Intelligente As Boolean) As StringBuilder
        Dim ope As New OperazioniSuFile

        If Intelligente Then
            ope.ScriveLogServizio(idProc, "SINCRONIZZAZIONE INTELLIGENTE DIRECTORY: " & Origine & "->" & Destinazione)
        Else
            ope.ScriveLogServizio(idProc, "SINCRONIZZAZIONE DIRECTORY: " & Origine & "->" & Destinazione)
        End If

        log = New StringBuilder

        ' Try
        Dim Rec2 As Object = CreateObject("ADODB.Recordset")
        Dim Sql As String
        Dim CartelleDest() As String = {}
        Dim qCartelleDest As Integer
        Dim LeggiFiles As Boolean

        Dim Filetti() As String = {}
        Dim DimensioneFiletti() As Long = {}
        Dim DataFiletti() As Date = {}
        Dim qFiletti As Long
        Dim CartelleOrig() As String = {}
        Dim qCartelleOrig As Long
        Dim Altro As String
        Dim Datella As Date
        Dim DatellaFile As String
        Dim Massimo As Long = 0
        Dim AggiornataTabellaIntelligente As Boolean

        ' Lettura origine
        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Lettura directory " & Origine, " ")

        If MetteInPausa Then
            MetteInPausaLaRoutine()
        End If

        If Not BloccaTutto Then
            gf.ScansionaDirectorySingola(Origine, Filtro, lblOperazione, False)

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            If Not BloccaTutto Then
                Filetti = gf.RitornaFilesRilevati
                qFiletti = gf.RitornaQuantiFilesRilevati

                CartelleOrig = gf.RitornaDirectoryRilevate
                qCartelleOrig = gf.RitornaQuanteDirectoryRilevate
            End If
        End If

        If MetteInPausa Then
            MetteInPausaLaRoutine()
        End If

        If Not BloccaTutto Then
            Sql = "Delete From FilesOrigine"
            Db.EsegueSql(idProc, ConnSQL, Sql)

            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Scrittura dati directory " & Origine, " ")
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Numero Cartelle: " & qCartelleOrig & " - Numero files: " & qFiletti, " ")

            If Mid(Origine, Origine.Length, 1) = "\" Then
                Altro = ""
            Else
                Altro = "\"
            End If

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    If i / 500 = Int(i / 500) Then
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "        " & i & "/" & qFiletti, " ")
                    End If

                    Try
                        Datella = FileDateTime(Filetti(i))
                        DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                    Catch ex As Exception
                        DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                    End Try

                    Sql = "Insert Into FilesOrigine Values (" &
                            "'" & Filetti(i).Replace("'", "''").Replace(Origine & Altro, "") & "', " &
                            " " & FileLen(Filetti(i)) & ", " &
                            "'" & DatellaFile & "' " &
                            ")"
                    Db.EsegueSql(idProc, ConnSQL, Sql)

                    If IsNothing(lblContatore) = False Then
                        lblContatore.Text = i & "/" & qFiletti
                    End If

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    If BloccaTutto Then
                        Exit For
                    End If

                    Application.DoEvents()
                End If
            Next

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            ReDim Filetti(0)
            ReDim DimensioneFiletti(0)
            ReDim DataFiletti(0)

            If Not BloccaTutto Then
                ' Lettura Destinazione
                LeggiFiles = True

                If Intelligente Then
                    ' Nel caso di sincronia intelligente vado a riprendere gli eventuali dati salvati nel db l'ultima volta. Se non ci sono li ricreo
                    LeggiFiles = False

                    Sql = "Select Max(Progressivo)+1 From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                    Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                    If Rec2(0).Value Is DBNull.Value = True Then
                        Massimo = 1
                    Else
                        Massimo = Rec2(0).Value
                    End If
                    Rec2.Close()

                    Sql = "Select Count(*) From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                    Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                    If Rec2(0).Value Is DBNull.Value = True Then
                        LeggiFiles = True
                    Else
                        If Rec2(0).Value = 0 Then
                            LeggiFiles = True
                        End If
                    End If
                    Rec2.Close()
                End If

                If LeggiFiles Then
                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Lettura directory " & Destinazione, " ")

                    gf.ScansionaDirectorySingola(Destinazione, Filtro, lblOperazione, False)

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    If Not BloccaTutto Then
                        Filetti = gf.RitornaFilesRilevati
                        DimensioneFiletti = gf.RitornaDimensioneFilesRilevati
                        DataFiletti = gf.RitornaDataFilesRilevati
                        qFiletti = gf.RitornaQuantiFilesRilevati

                        CartelleDest = gf.RitornaDirectoryRilevate
                        qCartelleDest = gf.RitornaQuanteDirectoryRilevate

                        If Intelligente Then
                            Sql = "Delete From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                        Else
                            Sql = "Delete From FilesDestinazione"
                        End If
                        Db.EsegueSql(idProc, ConnSQL, Sql)

                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Scrittura dati directory " & Destinazione, " ")
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Numero Cartelle: " & qCartelleDest & " - Numero files: " & qFiletti, " ")

                        If Mid(Destinazione, Destinazione.Length, 1) = "\" Then
                            Altro = ""
                        Else
                            Altro = "\"
                        End If

                        For i As Long = 0 To qFiletti
                            If Filetti(i) <> "" Then
                                If i / 500 = Int(i / 500) Then
                                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "        " & i & "/" & qFiletti, " ")
                                End If

                                Try
                                    Datella = FileDateTime(Filetti(i))
                                    DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                Catch ex As Exception
                                    DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                End Try

                                If Intelligente Then
                                    Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                        " " & idProc & ", " &
                                        " " & Progressivo & ", " &
                                        " " & (i + 1) & ", " &
                                        "'" & Filetti(i).Replace("'", "''").Replace(Destinazione & "\", "") & "', " &
                                        " " & DimensioneFiletti(i) & ", " &
                                        "'" & DatellaFile & "' " &
                                        ")"
                                Else
                                    Sql = "Insert Into FilesDestinazione Values (" &
                                        "'" & Filetti(i).Replace("'", "''").Replace(Destinazione & "\", "") & "', " &
                                        " " & DimensioneFiletti(i) & ", " &
                                        "'" & DatellaFile & "' " &
                                        ")"
                                End If
                                Db.EsegueSql(idProc, ConnSQL, Sql)
                                Massimo = i + 1

                                If IsNothing(lblContatore) = False Then
                                    lblContatore.Text = i & "/" & qFiletti
                                End If

                                If MetteInPausa Then
                                    MetteInPausaLaRoutine()
                                End If

                                If BloccaTutto Then
                                    Exit For
                                End If

                                Application.DoEvents()
                            End If
                        Next
                    End If
                Else
                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Lettura directory " & Destinazione & " skippata: Sincronizzazione intelligente", " ")
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                ' Copia i dati della sicnronia intelligente nella tabella di appoggio per rendere più veloci le ricerche
                If Not BloccaTutto Then
                    If Intelligente Then
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Pulizia valori per sincronizzazione intelligente", " ")
                        Sql = "Delete From DatiSincroniaIntelligente"
                        Db.EsegueSql(idProc, ConnSQL, Sql)

                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Copia valori per sincronizzazione intelligente", " ")
                        Sql = "Insert Into DatiSincroniaIntelligente Select * From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                        Db.EsegueSql(idProc, ConnSQL, Sql)
                    End If
                End If

                Dim FilesDaElaborare As Collection
                Dim Dimens As Long

                If Not BloccaTutto Then
                    ' Copia i file che esistono nell'origine ma non nella destinazione
                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Conteggio files da copiare verso destinazione", " ")

                    If Intelligente Then
                        Sql = "Select A.[File] From (FilesOrigine A Left Join DatiSincroniaIntelligente B " &
                            "On A.[File]=B.[File]) " &
                            "Where B.[File] Is Null And B.idProc=" & idProc & " And B.Operazione=" & Progressivo & ""
                    Else
                        Sql = "SELECT A.[File] " &
                            "FROM (FilesOrigine AS A LEFT OUTER JOIN " &
                            "FilesDestinazione AS B ON B.[File] = A.[File]) " &
                            "WHERE B.[File] Is NULL"
                    End If
                    Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                    FilesDaElaborare = New Collection
                    Do Until Rec2.Eof
                        FileDestinazione = Destinazione & "\" & Rec2("File").Value
                        If Not File.Exists(FileDestinazione) Then
                            FilesDaElaborare.Add(Rec2("File").Value)

                            Application.DoEvents()
                        End If

                        Rec2.MoveNext()
                    Loop
                    Rec2.Close()

                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Files rilevati: " & FilesDaElaborare.Count, " ")

                    If FilesDaElaborare.Count > 0 Then
                        Rec2.MoveFirst()

                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Elaborazione files da copiare verso destinazione", " ")
                        For i As Long = 1 To FilesDaElaborare.Count
                            FileOrigine = Origine & "\" & FilesDaElaborare.Item(i)
                            FileDestinazione = Destinazione & "\" & FilesDaElaborare.Item(i)

                            If File.Exists(FileOrigine) Then
                                gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(FileDestinazione) & "\")
                                If gf.CopiaFileFisico(FileOrigine, FileDestinazione, True) <> "SKIPPED" Then
                                    If Intelligente Then
                                        Try
                                            Datella = FileDateTime(FileOrigine)
                                            DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                        Catch ex As Exception
                                            DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                        End Try

                                        Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                            " " & idProc & ", " &
                                            " " & Progressivo & ", " &
                                            " " & Massimo & ", " &
                                            "'" & FileOrigine.Replace(Origine & "\", "").Replace("'", "''") & "', " &
                                            " " & FileLen(FileOrigine) & ", " &
                                            "'" & DatellaFile & "' " &
                                            ")"
                                        Db.EsegueSql(idProc, ConnSQL, Sql)

                                        Massimo += 1
                                    End If

                                    Dimens = Int(FileLen(FileOrigine) / 1024)
                                    If Dimens = 0 Then
                                        Dimens = FileLen(FileOrigine)
                                    End If
                                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Copia file " & FileOrigine & " (" & Int(FileLen(FileOrigine) / 1024) & " Kb.)", i & "/" & FilesDaElaborare.Count)
                                Else
                                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Skip file " & FileOrigine, i & "/" & FilesDaElaborare.Count)
                                End If
                            Else
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "File di origine non presente: " & FileOrigine, i & "/" & FilesDaElaborare.Count)
                            End If

                            If MetteInPausa Then
                                MetteInPausaLaRoutine()
                            End If

                            If BloccaTutto Then
                                Exit For
                            End If

                            'Rec2.MoveNext()
                        Next
                    End If
                    'Rec2.Close()

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    If Not BloccaTutto Then
                        ' Elimina i files nella destinazione che non esistono nell'origine
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Conteggio files da eliminare nella destinazione", " ")

                        If Intelligente Then
                            Sql = "Select A.[File] From (FileDestinazioneIntelligente A " &
                                "Left Join Filesorigine B On A.[File]=B.[File]) " &
                                "Where B.[File] Is Null And A.idProc = " & idProc & " And A.Operazione = " & Progressivo
                        Else
                            Sql = "SELECT  A.[File] " &
                                "FROM (FilesDestinazione AS A LEFT OUTER JOIN " &
                                "FilesOrigine AS B ON B.[File] = A.[File]) " &
                                "WHERE B.[File] Is NULL"
                        End If
                        Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                        FilesDaElaborare = New Collection
                        Do Until Rec2.Eof
                            FileOrigine = Origine & "\" & Rec2("File").Value
                            FileDestinazione = Destinazione & "\" & Rec2("File").Value
                            If Not File.Exists(FileOrigine) And File.Exists(FileDestinazione) Then
                                FilesDaElaborare.Add(Rec2("File").Value)

                                Application.DoEvents()
                            End If

                            Rec2.MoveNext()
                        Loop
                        Rec2.Close()

                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Files rilevati: " & FilesDaElaborare.Count, " ")

                        If FilesDaElaborare.Count > 0 Then
                            If FilesDaElaborare.Count <= 1000 Then
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Elaborazione files da eliminare nella destinazione", " ")
                                For i As Long = 1 To FilesDaElaborare.Count
                                    FileOrigine = Destinazione & "\" & FilesDaElaborare.Item(i)

                                    'ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione file " & FileOrigine, i & "/" & FilesDaElaborare.Count)

                                    If Not gf.EliminaFileFisico(FileOrigine).Contains("ERRORE:") Then
                                        If Intelligente Then
                                            Sql = "Delete From FileDestinazioneIntelligente Where " &
                                                "[File]='" & FileOrigine.Replace(Destinazione & "\", "").Replace("'", "''") & "' And " &
                                                "idProc=" & idProc & " And " &
                                                "Operazione=" & Progressivo
                                            Db.EsegueSql(idProc, ConnSQL, Sql)
                                        End If

                                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione file " & FileOrigine, i & "/" & FilesDaElaborare.Count)
                                    End If

                                    If MetteInPausa Then
                                        MetteInPausaLaRoutine()
                                    End If

                                    If BloccaTutto Then
                                        Exit For
                                    End If

                                    'Rec2.MoveNext()
                                Next
                            Else
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione file: Troppi files. Skip", FilesDaElaborare.Count)
                            End If
                        End If
                        'Rec2.Close()

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If Not BloccaTutto Then
                            ' Copia nella destinazione i files che hanno date superiori nell'origine o dimensioni diverse
                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Conteggio files diversi", " ")

                            If Intelligente Then
                                Sql = "Select A.[File] From (DatiSincroniaIntelligente A Left Join " &
                                    "FilesOrigine B " &
                                    "On A.[File]=B.[File]) " &
                                    "Where (A.Dimensione <> B.Dimensioni Or A.DataFile < B.DataOra) And A.idProc=" & idProc & " And A.Operazione=" & Progressivo & ""
                            Else
                                Sql = "Select A.[File] From (Filesorigine A Left Join FilesDestinazione B On A.[File] = B.[File]) " &
                                    "Where A.Dimensioni <> B.Dimensioni " &
                                    "Or A.DataOra > B.DataOra"
                            End If
                            Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                            FilesDaElaborare = New Collection
                            Do Until Rec2.Eof
                                FileOrigine = Origine & "\" & Rec2("File").Value
                                FileDestinazione = Destinazione & "\" & Rec2("File").Value
                                If File.Exists(FileOrigine) And File.Exists(FileDestinazione) Then
                                    FilesDaElaborare.Add(Rec2("File").Value)

                                    Application.DoEvents()
                                End If

                                Rec2.MoveNext()
                            Loop
                            Rec2.Close()

                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Files rilevati: " & FilesDaElaborare.Count, " ")

                            If FilesDaElaborare.Count > 0 Then
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Elaborazione files diversi", " ")

                                For i As Long = 1 To FilesDaElaborare.Count
                                    FileOrigine = Origine & "\" & FilesDaElaborare.Item(i)
                                    FileDestinazione = Destinazione & "\" & FilesDaElaborare.Item(i)

                                    If File.Exists(FileOrigine) Then
                                        gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(FileDestinazione) & "\")
                                        If gf.CopiaFileFisico(FileOrigine, FileDestinazione, True) <> "SKIPPED" Then
                                            If Intelligente Then
                                                Try
                                                    Datella = FileDateTime(FileOrigine)
                                                    DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                                Catch ex As Exception
                                                    DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                                End Try

                                                Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                                    " " & idProc & ", " &
                                                    " " & Progressivo & ", " &
                                                    " " & Massimo & ", " &
                                                    "'" & FileOrigine.Replace(Destinazione & "\", "").Replace("'", "''") & "', " &
                                                    " " & FileLen(FileOrigine) & ", " &
                                                    "'" & DatellaFile & "' " &
                                                    ")"
                                                Db.EsegueSql(idProc, ConnSQL, Sql)

                                                Massimo += 1
                                            End If

                                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Aggiorna file " & FileOrigine & " (" & Int(FileLen(FileOrigine) / 1024) & " Kb.)", i & "/" & FilesDaElaborare.Count)
                                        Else
                                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Skip file " & FileOrigine, i & "/" & FilesDaElaborare.Count)
                                        End If
                                    Else
                                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "File di origine non presente: " & FileOrigine, i & "/" & FilesDaElaborare.Count)
                                    End If

                                    If MetteInPausa Then
                                        MetteInPausaLaRoutine()
                                    End If

                                    If BloccaTutto Then
                                        Exit For
                                    End If

                                    'Rec2.MoveNext()
                                Next
                            End If
                            'Rec2.Close()

                            If MetteInPausa Then
                                MetteInPausaLaRoutine()
                            End If

                            If Not BloccaTutto Then
                                ' Elimina cartelle vuote nella destinazione
                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Pulizia tabella appoggio origine", "")
                                Sql = "Delete From DirectOrig"
                                Db.EsegueSql(idProc, ConnSQL, Sql)

                                Dim i As Long = 0

                                For Each cd As String In CartelleOrig
                                    If cd <> "" Then
                                        If i / 100 = Int(i / 100) Then
                                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Scrittura tabella origine", i & "/" & qCartelleOrig)
                                        End If
                                        cd = cd.Replace(Origine, "")
                                        If cd <> "" Then
                                            Sql = "Insert Into DirectOrig Values ('" & cd.Replace("'", "''") & "')"
                                            Db.EsegueSql(idProc, ConnSQL, Sql)
                                        End If
                                    End If
                                    i += 1
                                Next

                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Pulizia tabella appoggio destinazione", "")
                                Sql = "Delete From DirectDest"
                                Db.EsegueSql(idProc, ConnSQL, Sql)

                                i = 0
                                For Each cd As String In CartelleDest
                                    If cd <> "" Then
                                        If i / 100 = Int(i / 100) Then
                                            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Scrittura tabella destin.", i & "/" & qCartelleDest)
                                        End If
                                        cd = cd.Replace(Destinazione, "")
                                        If cd <> "" Then
                                            Sql = "Insert Into DirectDest Values ('" & cd.Replace("'", "''") & "')"
                                            Db.EsegueSql(idProc, ConnSQL, Sql)
                                        End If
                                    End If
                                    i += 1
                                Next

                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Lettura directory da eliminare", "")
                                Sql = "Select * From DirectDest Where Nome Not In (Select Nome From DirectOrig)"
                                Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                                Dim CartelleDaEliminare As New ArrayList
                                i = 0
                                Do Until Rec2.Eof
                                    If i / 100 = Int(i / 100) Then
                                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione tabella" & Destinazione & Rec2("Nome").Value, i & "/" & CartelleDaEliminare.Count)
                                    End If
                                    i += 1
                                    CartelleDaEliminare.Add(Destinazione & Rec2("Nome").Value)
                                    Rec2.MoveNext
                                Loop
                                Rec2.Close

                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Lettura directory da creare", "")
                                Sql = "Select * From DirectOrig Where Nome Not In (Select Nome From DirectDest)"
                                Rec2 = Db.LeggeQuery(idProc, ConnSQL, Sql)
                                Dim CartelleDaAggiungere As New ArrayList
                                i = 0
                                Do Until Rec2.Eof
                                    If i / 100 = Int(i / 100) Then
                                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Creazione tabella" & Destinazione & Rec2("Nome").Value, i & "/" & CartelleDaEliminare.Count)
                                    End If
                                    i += 1
                                    CartelleDaAggiungere.Add(Destinazione & Rec2("Nome").Value)
                                    Rec2.MoveNext
                                Loop
                                Rec2.Close

                                i = 0
                                Dim qc As Long = CartelleDaEliminare.Count
                                For Each c As String In CartelleDaEliminare
                                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione cartella: " & vbCrLf & c, i & "/" & qc)
                                    Directory.Delete(c, True)
                                Next

                                i = 0
                                qc = CartelleDaAggiungere.Count
                                For Each c As String In CartelleDaAggiungere
                                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Eliminazione cartella: " & vbCrLf & c, i & "/" & qc)
                                    Directory.CreateDirectory(c)
                                Next

                                'For k As Integer = 1 To 3
                                '    If Not CartelleDest Is Nothing Then
                                '        For i As Long = qCartelleDest - 1 To 0 Step -1
                                '            If CartelleDest(i) <> "" Then
                                '                If Directory.Exists(CartelleDest(i)) = True Then
                                '                    Dim Ok As Boolean = True

                                '                    If Not Cartelle Is Nothing Then
                                '                        For kk As Long = 1 To qCartelle - 1
                                '                            If Cartelle(kk).Replace(Origine & "\", "").Trim.ToUpper = CartelleDest(i).Replace(Destinazione & "\", "").Trim.ToUpper Then
                                '                                Ok = False
                                '                                Exit For
                                '                            End If
                                '                        Next
                                '                    End If

                                '                    If Ok Then
                                '                        Dim FilesInFolder As Integer = Directory.GetFiles(CartelleDest(i), "*.*").Count
                                '                        Dim FoldersInFolder As Integer = Directory.GetDirectories(CartelleDest(i), "*.*").Count

                                '                        If FilesInFolder = 0 And FoldersInFolder = 0 Then
                                '                            Try
                                '                                RmDir(CartelleDest(i))
                                '                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "      " & CartelleDest(i), " ")
                                '                            Catch ex As Exception

                                '                            End Try
                                '                        End If
                                '                    End If
                                '                End If
                                '            End If

                                '            If MetteInPausa Then
                                '                MetteInPausaLaRoutine()
                                '            End If

                                '            If BloccaTutto Then
                                '                Exit For
                                '            End If
                                '        Next

                                '        If MetteInPausa Then
                                '            MetteInPausaLaRoutine()
                                '        End If

                                '        If BloccaTutto Then
                                '            Exit For
                                '        End If
                                '    End If
                                'Next

                                If MetteInPausa Then
                                    MetteInPausaLaRoutine()
                                End If

                                If Not BloccaTutto Then
                                    '' Crea Cartelle nella destinazione in caso ce ne siano di vuote nell'origine
                                    'Dim sCartella As String

                                    'ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Creazione cartelle vuote presenti su origine", " ")

                                    'For i As Long = 0 To qCartelle - 1
                                    '    If Cartelle(i) <> "" Then
                                    '        sCartella = Cartelle(i).Replace(Origine, "")
                                    '        If sCartella <> "" Then
                                    '            If Mid(sCartella, 1, 1) = "\" Then
                                    '                sCartella = Mid(sCartella, 2, sCartella.Length)
                                    '            End If
                                    '            If Directory.Exists(Destinazione & "\" & sCartella) = False Then
                                    '                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "      " & Destinazione & "\" & sCartella & "\", " ")

                                    '                gf.CreaDirectoryDaPercorso(Destinazione & "\" & sCartella & "\")
                                    '            End If
                                    '        End If
                                    '    End If

                                    '    If MetteInPausa Then
                                    '        MetteInPausaLaRoutine()
                                    '    End If

                                    '    If BloccaTutto Then
                                    '        Exit For
                                    '    End If
                                    'Next

                                    If MetteInPausa Then
                                        MetteInPausaLaRoutine()
                                    End If

                                    If Not BloccaTutto Then
                                        ' Pulizia tabelle di appoggio e compattazione DB
                                        If Intelligente Then
                                            AggiornataTabellaIntelligente = True
                                            AggiornaTabellaIntelligente(idProc, Progressivo, Db, ConnSQL, lblOperazione, lblContatore)
                                        Else
                                            If MetteInPausa Then
                                                MetteInPausaLaRoutine()
                                            End If

                                            If Not BloccaTutto Then
                                                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Pulizia tabelle e compattazione DB", " ")

                                                Sql = "Delete From FilesOrigine"
                                                Db.EsegueSql(idProc, ConnSQL, Sql)

                                                Sql = "Delete From FilesDestinazione"
                                                Db.EsegueSql(idProc, ConnSQL, Sql)
                                            End If
                                        End If

                                        If MetteInPausa Then
                                            MetteInPausaLaRoutine()
                                        End If

                                        If Not BloccaTutto Then
                                            Db.CompattazioneDb()
                                            ' Pulizia tabelle di appoggio e compattazione DB

                                            If Intelligente Then
                                                ope.ScriveLogServizio(idProc, "USCITA SINCRONIZZAZIONE INTELLIGENTE")
                                            Else
                                                ope.ScriveLogServizio(idProc, "USCITA SINCRONIZZAZIONE")
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If BloccaTutto Then
            ope.ScriveLogServizio(idProc, "SINCRONIZZAZIONE BLOCCATA")
        Else
            If Intelligente Then
                If AggiornataTabellaIntelligente = False Then
                    AggiornaTabellaIntelligente(idProc, Progressivo, Db, ConnSQL, lblOperazione, lblContatore)
                End If
            End If
        End If
        'Catch ex As Exception
        '    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        'End Try
        ope = Nothing

        Return log
    End Function

    Private Sub AggiornaTabellaIntelligente(idProc As Integer, Progressivo As Integer, DB As GestioneACCESS, ConnSql As Object, lblOperazione As Label, lblContatore As Label)
        Dim Sql As String = ""

        ' Copio la tabella di origine in quella di destinazione in caso di sincronia intelligente
        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Aggiornamento tabella di sincronia", " ")

        Sql = "Delete From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
        DB.EsegueSql(idProc, ConnSql, Sql)

        Dim c As Long = 0
        Dim d As Date
        Dim mese As String
        Dim giorno As String
        Dim ora As String
        Dim minuti As String
        Dim secondi As String
        Dim Rec2 As Object = CreateObject("ADODB.Recordset")
        Dim DatellaFile As String = ""

        Sql = "Select * From FilesOrigine"
        Rec2 = DB.LeggeQuery(idProc, ConnSql, Sql)
        Do Until Rec2.Eof
            c += 1
            d = Rec2(2).Value

            mese = d.Month.ToString.Trim : If mese.Length = 1 Then mese = "0" & mese
            giorno = d.Day.ToString.Trim : If giorno.Length = 1 Then giorno = "0" & giorno
            ora = d.Hour.ToString.Trim : If ora.Length = 1 Then ora = "0" & ora
            minuti = d.Minute.ToString.Trim : If minuti.Length = 1 Then minuti = "0" & minuti
            secondi = d.Second.ToString.Trim : If secondi.Length = 1 Then secondi = "0" & secondi

            DatellaFile = d.Year & "-"
            DatellaFile &= mese & "-"
            DatellaFile &= giorno & " "
            DatellaFile &= ora & ":"
            DatellaFile &= minuti & ":"
            DatellaFile &= secondi

            Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                 " " & idProc & ", " &
                 " " & Progressivo & ", " &
                 " " & c & ", " &
                 "'" & Replace(Rec2(0).Value, "'", "''") & "', " &
                 " " & Rec2(1).Value & ", " &
                 "'" & DatellaFile & "' " &
                 ")"
            DB.EsegueSql(idProc, ConnSql, Sql)

            Rec2.MoveNext()
        Loop
        Rec2.Close()

        Sql = "Delete From FilesOrigine"
        DB.EsegueSql(idProc, ConnSql, Sql)

        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Aggiornate righe sincronia: " & c & " - idProc: " & idProc & " - Operazione: " & Progressivo, " ")
    End Sub

    Public Function EsegueZip(idProc As Integer, Origine As String, Destinazione As String, lblOperazione As Label, lblContatore As Label, ModalitaServizo As Boolean) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "ZIP FILES: " & Origine & "->" & Destinazione)

        Dim PathUlteriore As String = ""

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, "", lblOperazione, False)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati

            Try
                File.Delete(Destinazione)
            Catch ex As Exception

            End Try

            Try
                Using zip As New ZipFile()
                    Dim FileZippante As ZipEntry

                    zip.ParallelDeflateThreshold = -1

                    ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Aggiunta file a zip", " ")

                    For i As Long = 1 To qFiletti
                        FileZippante = zip.AddFile(Filetti(i))

                        If BloccaTutto Then
                            Exit For
                        End If
                    Next

                    If Not BloccaTutto Then
                        ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Salvataggio zip", " ")

                        zip.Save(Destinazione)
                    End If
                End Using
            Catch ex1 As System.Exception
                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Errore su zip: " + ex1.Message, " ")
            End Try
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

    Public Function ListaFiles(idProc As Integer, Origine As String, Destinazione As String, lblOperazione As Label, lblContatore As Label, ModalitaServizio As Boolean) As StringBuilder
        Dim ope As New OperazioniSuFile
        ope.ScriveLogServizio(idProc, "LISTA FILES: " & Origine & "->" & Destinazione)

        Dim PathUlteriore As String = ""

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, "", lblOperazione, False)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati
            Dim qDirectory As Long = gf.RitornaQuanteDirectoryRilevate
            Dim Percorso As String
            Dim oldPercorso As String = ""
            Dim Filetto As String
            Dim qDett As Integer = 0
            Dim qBytes As Long = 0
            Dim qBytesTotali As Long = 0

            Try
                File.Delete(Destinazione)
            Catch ex As Exception

            End Try

            Try
                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Creazione lista files", " ")

                gf.ApreFileDiTestoPerScrittura(Destinazione)
                For i As Long = 1 To qFiletti
                    Percorso = gf.TornaNomeDirectoryDaPath(Filetti(i))
                    Filetto = gf.TornaNomeFileDaPath(Filetti(i))

                    If oldPercorso <> Percorso Then
                        If oldPercorso <> "" Then
                            gf.ScriveTestoSuFileAperto("")
                            gf.ScriveTestoSuFileAperto("     Files: " & gf.FormattaNumero(qDett, False) & " - MBytes: " & gf.FormattaNumero(qBytes / 1024 / 1024, False))
                            gf.ScriveTestoSuFileAperto("")
                        End If
                        gf.ScriveTestoSuFileAperto(Percorso)
                        gf.ScriveTestoSuFileAperto("")
                        oldPercorso = Percorso
                        qDett = 0
                        qBytes = 0
                    End If

                    gf.ScriveTestoSuFileAperto("     " & Filetto)
                    qDett += 1
                    qBytes += gf.TornaDimensioneFile(Filetti(i))
                    qBytesTotali += gf.TornaDimensioneFile(Filetti(i))

                    If BloccaTutto Then
                        Exit For
                    End If
                Next
                gf.ScriveTestoSuFileAperto("")
                gf.ScriveTestoSuFileAperto("     Directories: " & gf.FormattaNumero(qDirectory, False) & " - Files: " & gf.FormattaNumero(qFiletti, False) & " - MBytes: " & gf.FormattaNumero(qBytesTotali / 1024 / 1024, False))
                gf.ChiudeFileDiTestoDopoScrittura()
            Catch ex1 As System.Exception
                ScriveOperazione(idProc, log, lblOperazione, lblContatore, "Errore su Lista Files: " + ex1.Message, " ")
            End Try
        Catch ex As Exception
            ScriveOperazione(idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ")
        End Try
        ope = Nothing

        Return log
    End Function

End Class
