Imports System.Text
Imports System.Timers
Imports System.IO

Public Class BackupNET

    ' Usare come amministratore di sistema il
    ' Prompt dei comandi degli strumenti nativi di VS2012 x86
    '
    ' installutil Backup.NET.exe
    ' installutil /u Backup.NET.exe
    '
    ' RICORDARSI DI AGGIORNARE LE CLASSI OPERAZIONISUFILE.VB E OPERAZIONISUFILEDETTAGLI.VB in caso
    ' di modifica

    Private MinutoElaborazione As Integer = -1
    Private tmrMain As Timer
    Private idProceduraDaEseguire() As String = {}
    Private qId As Integer = 0

    Public Sub BackupNET()
        InitializeComponent()
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        'PulisceFileDiLog()

        Try
            MkDir("C:\BackupLog")
        Catch ex As Exception

        End Try

        tmrMain = New Timer
        AddHandler tmrMain.Elapsed, AddressOf tmrMain_Contatore
        AddHandler tmrMain.Disposed, AddressOf tmrMain_Disposed
        tmrMain.Interval = 30000
        tmrMain.Enabled = True

        Dim gf As New GestioneFilesDirectory
        Dim Config As String = ""
        If File.Exists("C:\BackupLog\Config\Config.dat") Then
            Config = gf.LeggeFileIntero("C:\BackupLog\Config\Config.dat")
        Else
            gf.CreaDirectoryDaPercorso("C:\BackupLog\Config\")
            gf.ApreFileDiTestoPerScrittura("C:\BackupLog\Config\Config.dat")
            Config = "C:\Program Files (x86)\Backup.NET\BackupMain.NET.exe;looigi@gmail.com;pippuzzo227!;"
            gf.ScriveTestoSuFileAperto(Config)
            gf.ChiudeFileDiTestoDopoScrittura()
        End If
        gf = Nothing

        Dim c() As String = Config.Split(";")
        PercorsoApplicazione = c(0)

        Dim opFiles As RoutineVarie = New RoutineVarie
        opFiles.ScriveLogServizio(0, "Servizio partito")
        opFiles.ScriveLogServizio(0, "Path applicazione: " & c(0))
        opFiles.ScriveLogServizio(0, "Utenza mail: " & c(1))
        opFiles.ScriveLogServizio(0, "Password mail: " & c(2))
        opFiles = Nothing
    End Sub

    Protected Overrides Sub OnStop()
        Dim opFiles As RoutineVarie = New RoutineVarie
        opFiles.ScriveLogServizio(0, "Servizio fermato")
        opFiles = Nothing
    End Sub

    Private Sub tmrMain_Contatore(ByVal source As Object, ByVal e As ElapsedEventArgs)
        tmrMain.Stop()
        GC.Collect()

        'If File.Exists(NomeFileLog) Then
        '    Dim d As Date = FileDateTime(NomeFileLog)
        '    Dim diff As Integer = Math.Abs(DateDiff(DateInterval.Day, Now, d))
        '    If diff > 30 Then
        '        Try
        '            Kill(NomeFileLog)
        '        Catch ex As Exception

        '        End Try
        '    End If
        'End If

        Dim opFiles As RoutineVarie = New RoutineVarie
        'opFiles.PulisceNomeFileLog()

        If MinutoElaborazione <> Now.Minute Then
            MinutoElaborazione = Now.Minute

            TornaOperazioniDaEseguire(opFiles)

            If Not idProceduraDaEseguire Is Nothing Then
                If idProceduraDaEseguire.Length - 1 > -1 Then
                    For i As Integer = 1 To idProceduraDaEseguire.Length - 1
                        If idProceduraDaEseguire(i) <> "" Then
                            opFiles.ScriveLogServizio(1, "")
                            opFiles.ScriveLogServizio(1, "-----------------------------------------------------------")
                            opFiles.ScriveLogServizio(1, "Eseguo procedura: " & idProceduraDaEseguire(i))

                            EsegueProcedura(idProceduraDaEseguire(i), opFiles)

                            opFiles.ScriveLogServizio(1, "-----------------------------------------------------------")
                            opFiles.ScriveLogServizio(1, "")
                        End If
                    Next
                End If
            End If
        End If
        opFiles = Nothing

        tmrMain.Start()
    End Sub

    Private Sub TornaOperazioniDaEseguire(opFiles As RoutineVarie)
        Try
            qId = 0
            Erase idProceduraDaEseguire

            Dim DB As GestioneACCESS = New GestioneACCESS

            If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
                Dim ConnSQL As Object = DB.ApreDB(0)

                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String

                ' Controlla se esistono backup per l'ora e il giorno della settimana
                Sql = "Select B.NomeProcedura From Schedulazioni A Left Join NomiProcedure B On A.idProc=B.idProc " &
                    "Where " &
                    "TipoBackup='G' And " &
                    "Instr(Valore1, '" & TornaNumeroGiornoSettimana() & ";')>0 And " &
                    "Orario='" & TornaOrario() & "'"

                'opFiles.ScriveLogServizio(idProc, "Controlla se esistono backup per l'ora e il giorno della settimana")
                'opFiles.ScriveLogServizio(idProc, Sql)

                Rec = DB.LeggeQuery(0, ConnSQL, Sql)
                Do Until Rec.Eof
                    qId += 1
                    ReDim Preserve idProceduraDaEseguire(qId)
                    idProceduraDaEseguire(qId) = Rec("NomeProcedura").Value

                    Rec.MoveNext()
                Loop
                Rec.Close()

                ' Controlla se esistono backup per l'ora e il giorno del mese
                Sql = "Select B.NomeProcedura From Schedulazioni A Left Join NomiProcedure B On A.idProc=B.idProc " &
                    "Where " &
                    "TipoBackup='M' And " &
                    "Instr(Valore1, '" & TornaNumeroGiornoMese() & ";')>0 And " &
                    "Orario='" & TornaOrario() & "'"

                'opFiles.ScriveLogServizio(idProc, "Controlla se esistono backup per l'ora e il giorno del mese")
                'opFiles.ScriveLogServizio(idProc, Sql)

                Rec = DB.LeggeQuery(0, ConnSQL, Sql)
                Do Until Rec.Eof
                    qId += 1
                    ReDim Preserve idProceduraDaEseguire(qId)
                    idProceduraDaEseguire(qId) = Rec("NomeProcedura").Value

                    Rec.MoveNext()
                Loop
                Rec.Close()

                ConnSQL.close()
                ConnSQL = Nothing
            Else
                opFiles.ScriveLogServizio(1, "ERRORE SU APERTURA DB")
            End If

            DB = Nothing
        Catch ex As Exception
            opFiles.ScriveLogServizio(1, "ERRORE PROCEDURALE: " & ex.Message)
        End Try
    End Sub

    Private Function TornaOrario() As String
        Dim Ora As String = Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00")

        Return Ora
    End Function

    Private Function TornaNumeroGiornoMese() As String
        Dim Giorno As String = Format(Now.Day, "00")

        Return Giorno
    End Function

    Private Function TornaNumeroGiornoSettimana() As String
        Dim numGiorno As Integer = Now.DayOfWeek
        Dim Ritorno As String = ""

        Select Case numGiorno
            Case DayOfWeek.Monday
                Ritorno = "1"
            Case DayOfWeek.Tuesday
                Ritorno = "2"
            Case DayOfWeek.Wednesday
                Ritorno = "3"
            Case DayOfWeek.Thursday
                Ritorno = "4"
            Case DayOfWeek.Friday
                Ritorno = "5"
            Case DayOfWeek.Saturday
                Ritorno = "6"
            Case DayOfWeek.Sunday
                Ritorno = "7"
        End Select

        Return Ritorno
    End Function

    Protected Sub tmrMain_Disposed(ByVal source As Object, ByVal e As EventArgs)
        Dim opFiles As RoutineVarie = New RoutineVarie
        opFiles.ScriveLogServizio(1, "Chiuso servizio")
        opFiles = Nothing
    End Sub

    Private Sub EsegueProcedura(NomeProcedura As String, opFiles As RoutineVarie)
        Dim p As Process = New Process()
        p.StartInfo.FileName = PercorsoApplicazione
        p.StartInfo.Arguments = "/RUN " & Chr(34) & NomeProcedura & Chr(34)
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.Start()

        Dim output As String = p.StandardOutput.ReadToEnd()
        p.WaitForExit()
    End Sub
End Class
