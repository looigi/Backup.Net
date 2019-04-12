Imports System.Text
Imports System.Timers
Imports System.Globalization
Imports System.ServiceProcess
Imports System.Windows.Forms
Imports System.IO
Imports System.Collections.Specialized

Module mdlGenerale
    Public PercorsoDB As String = ""
    Public EffettuaLog As Boolean = True
    Public log As StringBuilder = New StringBuilder
    Public MetteInPausa As Boolean = False
    Public BloccaTutto As Boolean = False
    Public PercorsoApplicazione As String = ""

    'Public NomeFileLog As String = ""

    Public Enum TipoDB
        Niente = -1
        Access = 1
        SQLCE = 2
    End Enum

    Public DBSql As Integer = TipoDB.Niente

    Public Enum TipoOperazione
        Nulla = 0
        Copia = 1
        Spostamento = 2
        Sincronizzazione = 3
        Eliminazione = 4
        CreaDirectory = 5
        EliminaDirectory = 6
        RiavvioPC = 7
        AvvioServizio = 8
        FermaServizio = 9
        AvviaEseguibile = 10
        FermaEseguibile = 11
        Attendi = 12
        SincroniaIntelligente = 13
        Zip = 14
        ListaFiles = 15
        Messaggio = 16
    End Enum

    Public Enum StrutturaTabella
        idProc = 0
        Progressivo = 1
        idOperazione = 2
        Origine = 3
        Destinazione = 4
        Sovrascrivi = 5
        Sottodirectory = 6
        Filtro = 7
        Parametro = 8
        UtenzaOrigine = 9
        PasswordOrigine = 10
        UtenzaDestinazione = 11
        PasswordDestinazione = 12
        Attivo = 13
    End Enum

    Public Sub ImpostaDBSqlAccess()
        Dim PercorsoDB As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", "")
        If Directory.Exists(PercorsoDB) = False Then
            PercorsoDB = ""
        End If
        If PercorsoDB = "" Then
            My.Computer.Registry.CurrentUser.CreateSubKey("Software\BackupNet")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", Application.StartupPath)
            PercorsoDB = Application.StartupPath
        End If

        If File.Exists(PercorsoDB & "\DB\dbBackup.sdf") Then
            DBSql = TipoDB.SQLCE
        Else
            DBSql = TipoDB.Access
        End If
    End Sub

    Public Function LeggeUtenzaPassword()
        Dim Ritorno As String = ""

        'Try
        '    PercorsoDB = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", "")
        'Catch ex As Exception
        '    opFiles.ScriveLogServizio(idProc, "Percorso DB non impostato: " & ex.Message)
        '    End
        'End Try

        Dim gf As New GestioneFilesDirectory
        If File.Exists(PercorsoDB & "\Credenziali.txt") Then
            Ritorno = gf.LeggeFileIntero(PercorsoDB & "\Credenziali.txt")
        Else
            Ritorno = "looigi@gmail.com;pippuzzo227!;"
            gf.ApreFileDiTestoPerScrittura(PercorsoDB & "\Credenziali.txt")
            gf.ScriveTestoSuFileAperto(Ritorno)
            gf.ChiudeFileDiTestoDopoScrittura()
        End If
        gf = Nothing

        Return Ritorno
    End Function

    Public Sub ScriveOperazione(idProc As Integer, ByRef log As StringBuilder, lblOperazione As Label, lblContatore As Label, Optional Primo As String = "", Optional Secondo As String = "")
        If Primo <> "" Then
            If lblOperazione Is Nothing = False Then
                lblOperazione.Text = Primo
            End If
        End If

        If Secondo <> "" Then
            If lblContatore Is Nothing = False Then
                lblContatore.Text = Secondo
            End If
        End If

        Dim opFiles As New RoutineVarie
        opFiles.ScriveLogServizio(idProc, Primo & ";" & Secondo & ";")
        opFiles = Nothing

        log.Append(PrendeDataOra() & ";" & Primo & ";" & Secondo & ";" & vbCrLf)
    End Sub

    Public Function FormattaNumero(Numero As Single, ConVirgola As Boolean, Optional Lunghezza As Integer = -1) As String
        Dim Ritorno As String
        Dim Formattazione As String

        Select Case ConVirgola
            Case True
                Formattazione = "0,000.00"
            Case False
                Formattazione = "0,000"
            Case Else
                Formattazione = "0"
        End Select

        Ritorno = Numero.ToString(Formattazione, CultureInfo.InvariantCulture)

        Do While Left(Ritorno, 1) = "0"
            Ritorno = Mid(Ritorno, 2, Ritorno.Length)
        Loop
        If ConVirgola = True Then
            If Left(Ritorno.Trim, 1) = "." Then
                Ritorno = "0" & Ritorno
            End If
        Else
            If Left(Ritorno.Trim, 1) = "," Then
                Ritorno = Mid(Ritorno, 2, Ritorno.Length)
                For i As Integer = 1 To Ritorno.Length
                    If Mid(Ritorno, i, 1) = "0" Then
                        Ritorno = Mid(Ritorno, 1, i - 1) & "*" & Mid(Ritorno, i + 1, Ritorno.Length)
                    Else
                        Exit For
                    End If
                Next
                Ritorno = Ritorno.Replace("*", "")
                If Ritorno = "" Then Ritorno = "0"
            End If
        End If

        Ritorno = Ritorno.Replace(",", "+")
        Ritorno = Ritorno.Replace(".", ",")
        Ritorno = Ritorno.Replace("+", ".")

        If Ritorno = ".000" Then
            Ritorno = "0"
        End If

        If Lunghezza <> -1 Then
            Dim Spazi As String = ""

            If Ritorno.Length < Lunghezza Then
                For i As Integer = 1 To Lunghezza - Ritorno.Length
                    Spazi += " "
                Next
                Ritorno = Spazi & Ritorno
            End If
        End If

        Return Ritorno
    End Function

    Public Sub MetteInPausaLaRoutine()
        While MetteInPausa
            Threading.Thread.Sleep(1000)
        End While
    End Sub

    Private Function PrendeDataOra() As String
        Dim Ritorno As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

        Return Ritorno
    End Function

    'Public Sub PulisceFileDiLog()
    '    Dim gf As New GestioneFilesDirectory
    '    Dim opFiles As New OperazioniSuFile
    '    gf.EliminaFileFisico(opFiles.RitornaNomeFileDiLog)
    '    opFiles = Nothing
    '    gf = Nothing
    'End Sub

    Public Function FermaServizio(idProc As Integer, Nome As String) As Boolean
        Dim Ok As Boolean = True

        Try
            Dim opFiles As New RoutineVarie
            opFiles.ScriveLogServizio(idProc, "FERMA SERVIZIO:")

            Dim myController As ServiceController

            myController = New ServiceController(Nome)
            If myController.CanStop Then
                Try
                    myController.Stop()
                    opFiles.ScriveLogServizio(idProc, "SERVIZIO FERMATO")
                Catch ex As Exception
                    Ok = False
                    opFiles.ScriveLogServizio(idProc, "ERRORE FERMA SERVIZIO: " & ex.Message)
                End Try
            Else
                Ok = False
                opFiles.ScriveLogServizio(idProc, "ERRORE: SERVIZIO NON STOPPABILE")
            End If
            opFiles = Nothing
        Catch ex As Exception

        End Try

        Return Ok
    End Function

    Public Function FaiPartireServizio(idProc As Integer, Nome As String, Parametro As String) As Boolean
        Dim opFiles As New RoutineVarie
        opFiles.ScriveLogServizio(idProc, "AVVIA SERVIZIO: " & Nome)

        Dim myController As ServiceController
        Dim Ok As Boolean = True

        myController = New ServiceController(Nome)
        Try
            myController.Start()
        Catch exp As Exception
            Ok = False
            opFiles.ScriveLogServizio(idProc, "ERRORE AVVIA SERVIZIO: " & exp.Message)
        Finally
            opFiles.ScriveLogServizio(idProc, "SERVIZIO AVVIATO")
        End Try
        opFiles = Nothing

        Return Ok
    End Function

    Public Sub Attendi(ByVal gapToWait As Integer)
        Dim o As Integer = Now.Second
        Dim Diff As Integer
        Dim Ancora As Boolean = True

        Do While Ancora = True
            Diff = (Now.Second - o)
            If Diff < 0 Then
                Diff = 60 - Math.Abs(Diff)
            End If

            If Diff > gapToWait Then
                Ancora = False
            Else
            End If

            If BloccaTutto Then
                Exit Do
            End If
        Loop
    End Sub

    Public Function FindNextAvailableDriveLetter() As String
        ' build a string collection representing the alphabet
        Dim alphabet As New StringCollection()

        Dim lowerBound As Integer = Convert.ToInt16("a"c)
        Dim upperBound As Integer = Convert.ToInt16("z"c)
        For i As Integer = lowerBound To upperBound - 1
            Dim driveLetter As Char = ChrW(i)
            alphabet.Add(driveLetter.ToString())
        Next

        ' get all current drives
        Dim drives As DriveInfo() = DriveInfo.GetDrives()
        alphabet.Remove("a")
        alphabet.Remove("b")
        For Each drive As DriveInfo In drives
            alphabet.Remove(drive.Name.Substring(0, 1).ToLower())
        Next

        If alphabet.Count > 0 Then
            Return alphabet(0).ToUpper
        Else
            Return ""
        End If
    End Function

    'Public Sub ScriveLog(idLog As Integer, Cosa As String)
    '    Try
    '        Dim gf As New GestioneFilesDirectory
    '        Dim opef As New OperazioniSuFile
    '        Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
    '        gf.ApreFileDiTestoPerScrittura(opef.RitornaNomeFileDiLog)
    '        gf.ScriveTestoSuFileAperto("Procedura: " & idLog & " " & Datella & " -> " & Cosa)
    '        gf.ChiudeFileDiTestoDopoScrittura()
    '        opef = Nothing
    '        gf = Nothing
    '    Catch ex As Exception

    '    End Try
    'End Sub
End Module
