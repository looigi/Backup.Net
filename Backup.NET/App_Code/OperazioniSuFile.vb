Imports System.Windows.Forms
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.ServiceProcess
Imports System.IO

Public Class OperazioniSuFile
    Private DB As New GestioneACCESS
    Private ConnSQL As Object

    Sub New()
        MetteInPausa = False
        BloccaTutto = False

        ImpostaDBSqlAccess()
    End Sub

    Public Sub ImpostaPausa(Attiva As Boolean)
        MetteInPausa = Attiva
    End Sub

    Public Sub ScriveLogServizio(idProc As Integer, Log As String)
        Try
            Dim gf As New GestioneFilesDirectory
            Dim opef As New OperazioniSuFile
            Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
            gf.ApreFileDiTestoPerScrittura("C:\BackupLog\Servizio.txt")
            gf.ScriveTestoSuFileAperto("Procedura: " & idProc & " " & Datella & " -> " & Log)
            gf.ChiudeFileDiTestoDopoScrittura()
            opef = Nothing
            gf = Nothing
        Catch ex As Exception

        End Try
    End Sub

    'Public Sub ScriveLogDaRemoto(idProc As Integer, Log As String)
    '    ScriveLog(idProc, Log)
    'End Sub

    Public Sub ImpostaBlocco()
        BloccaTutto = True
    End Sub

    Public Function TornaRecordsetRigheProcedura(idProc As Integer) As String(,)
        Dim Recc(,) As String = {}
        Dim opeFiles As New OperazioniSuFile

        opeFiles.ScriveLogServizio(idProc, "Torna recordset Righe Procedura")

        Try
            If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
                ConnSQL = DB.ApreDB(idProc)
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Riga As Integer = 0
                Dim Colonna As Integer = 0
                Dim Sql As String

                ReDim Preserve Recc(150, 13)

                Sql = "Select * From DettaglioProcedure Where idProc=" & idProc & " Order By Progressivo"
                Try
                    Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
                    Do Until Rec.eof
                        For i As Integer = 1 To 14
                            Recc(Riga, i - 1) = Rec(i - 1).Value
                        Next
                        Riga += 1

                        Rec.MoveNext()
                    Loop
                    Rec.Close()

                    Recc(Riga, 0) = "***"
                Catch ex As Exception
                    opeFiles.ScriveLogServizio(idProc, "Torna Recordset. Errore: " & ex.Message)
                    Stop
                End Try

                Dim Ok As String = "True"
                If Rec Is Nothing Then
                    Ok = "False"
                End If
                opeFiles.ScriveLogServizio(idProc, "Torna Recordset. Righe: " & Riga - 1 & " Colonne: " & Colonna)
            Else
                opeFiles.ScriveLogServizio(idProc, "Errore nell'apertura del recordset")
            End If
        Catch ex As Exception
            opeFiles.ScriveLogServizio(idProc, "ERRORE: " & ex.Message)
        End Try
        opeFiles = Nothing

        Return Recc
    End Function

    Public Sub RilasciaOggetti()
        ConnSQL.close()
        ConnSQL = Nothing

        DB = Nothing
    End Sub

    Private Function SistemaLunghezzaCampo(Campo As String, Lunghezza As String) As String
        Dim Ritorno As String = Campo.Trim

        If Ritorno.Length < Lunghezza Then
            For i As Integer = Ritorno.Length To Lunghezza
                Ritorno = Ritorno & " "
            Next
        Else
            Ritorno = Mid(Ritorno, 1, (Lunghezza / 2)) & "..." & Mid(Ritorno, Ritorno.Length - (Lunghezza / 2) + 3, Lunghezza)
        End If

        Return Ritorno
    End Function

    Public Function CaricaDatiProcedura(idProc As Integer) As String
        Dim Ritorno As String = ""
        Dim opFiles As New OperazioniSuFile

        opFiles.ScriveLogServizio(idProc, "Carica Dati Procedura")

        Try
            Dim InvioMail As String = "N"
            Dim DB As New GestioneACCESS

            If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
                Dim ConnSQL As Object = DB.ApreDB(idProc)
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String = "Select * From NomiProcedure Where idProc=" & idProc

                opFiles.ScriveLogServizio(idProc, "Carica dati procedura: " & Sql)

                Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
                If Not Rec Is Nothing Then
                    If Rec.Eof = False Then
                        InvioMail = Rec("InvioMail").Value
                    End If
                    Rec.Close()
                Else
                    InvioMail = True
                End If

                opFiles.ScriveLogServizio(idProc, "InvioMail: " & InvioMail)

                Ritorno = InvioMail & ";"
            End If

            opFiles = Nothing
            DB = Nothing
        Catch ex As Exception
            opFiles.ScriveLogServizio(idProc, "ERRORE: " & ex.Message)
        End Try

        Return Ritorno
    End Function

    Public Function CaricaRigheProcedura(Procedura As String, Optional lstOperazioni As ListBox = Nothing) As Integer
        Dim idProc As Integer
        Dim opeFiles As New OperazioniSuFile

        Try
            opeFiles.ScriveLogServizio(idProc, "Carica righe procedura")

            Dim DB As New GestioneACCESS

            If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
                Dim ConnSQL As Object = DB.ApreDB(idProc)
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String

                idProc = -1

                If lstOperazioni Is Nothing = False Then
                    lstOperazioni.Items.Clear()
                End If

                Sql = "Select idProc From NomiProcedure Where NomeProcedura='" & Procedura & "'"
                Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
                If Rec.Eof = False Then
                    idProc = Rec(0).Value
                End If
                Rec.Close()

                If idProc <> -1 Then
                    Dim sTipoOperazione As String = ""
                    Dim sOrigine As String = ""
                    Dim sDestinazione As String = ""
                    Dim sSovrascrivi As String = ""
                    Dim sSottoDirectory As String = ""
                    Dim sFiltro As String = ""
                    Dim sProgressivo As String = ""

                    Sql = "Select * From DettaglioProcedure Where idProc=" & idProc & " Order By Progressivo"
                    Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
                    Do Until Rec.Eof
                        sProgressivo = Rec("Progressivo").Value
                        If sProgressivo.Length = 1 Then
                            sProgressivo = " " & sProgressivo
                        End If

                        Select Case Rec("idOperazione").Value
                            Case TipoOperazione.Nulla
                                sTipoOperazione = ""
                            Case TipoOperazione.Copia
                                sTipoOperazione = "Copia"
                            Case TipoOperazione.CreaDirectory
                                sTipoOperazione = "Crea dir"
                            Case TipoOperazione.EliminaDirectory
                                sTipoOperazione = "Elimina dir"
                            Case TipoOperazione.Eliminazione
                                sTipoOperazione = "Elimina Files"
                            Case TipoOperazione.Sincronizzazione
                                sTipoOperazione = "Sincronizza"
                            Case TipoOperazione.SincroniaIntelligente
                                sTipoOperazione = "Sincronia Intelligente"
                            Case TipoOperazione.Spostamento
                                sTipoOperazione = "Sposta"
                            Case TipoOperazione.RiavvioPC
                                sTipoOperazione = "Riavvio"
                            Case TipoOperazione.AvvioServizio
                                sTipoOperazione = "Avvia Servizio"
                            Case TipoOperazione.FermaServizio
                                sTipoOperazione = "Ferma Servizio"
                            Case TipoOperazione.AvviaEseguibile
                                sTipoOperazione = "Avvia EXE"
                            Case TipoOperazione.FermaEseguibile
                                sTipoOperazione = "Ferma EXE"
                            Case TipoOperazione.Attendi
                                sTipoOperazione = "Attendi"
                            Case TipoOperazione.Zip
                                sTipoOperazione = "Zip"
                            Case TipoOperazione.ListaFiles
                                sTipoOperazione = "Lista Files"
                            Case TipoOperazione.Messaggio
                                sTipoOperazione = "Messaggio"
                        End Select

                        For i As Integer = sTipoOperazione.Length To 23
                            sTipoOperazione = sTipoOperazione & " "
                        Next

                        sOrigine = SistemaLunghezzaCampo(Rec("Origine").Value, 40)
                        sDestinazione = SistemaLunghezzaCampo(Rec("Destinazione").Value, 40)

                        sSottoDirectory = Rec("Sottodirectory").Value
                        sSovrascrivi = Rec("Sovrascrivi").Value

                        sFiltro = SistemaLunghezzaCampo(Rec("Filtro").Value, 6)

                        opeFiles.ScriveLogServizio(idProc, sTipoOperazione & " " & sOrigine & " - " & sDestinazione & " " & sSottoDirectory & " " & sSovrascrivi & " " & sFiltro)

                        If lstOperazioni Is Nothing = False Then
                            lstOperazioni.Items.Add(sProgressivo & " " & sTipoOperazione & " " & sOrigine & " " & sDestinazione & " " & sSottoDirectory & " " & sSovrascrivi & " " & sFiltro)
                        End If

                        Rec.MoveNext()
                    Loop
                    Rec.Close()
                End If

                ConnSQL.close()
                ConnSQL = Nothing
            End If

            DB = Nothing
        Catch ex As Exception
            opeFiles.ScriveLogServizio(idProc, "ERRORE: " & ex.Message)
            idProc = -1
        End Try

        Return idProc
    End Function

    Private Function ProvaACreareFile(Lettera As String) As String
        Dim Ritorno As String = ""

        If Lettera <> "C:" And Lettera.Trim <> "" Then
            Dim Path As String = Lettera & "\Buttami.txt"

            Try
                Dim fs As FileStream = File.Create(Path)

                ' Add text to the file.
                Dim info As Byte() = New UTF8Encoding(True).GetBytes("ppp")
                fs.Write(info, 0, info.Length)
                fs.Close()
            Catch ex As Exception
                Ritorno = ex.Message
            End Try

            If Ritorno = "" Then
                Dim gf As New GestioneFilesDirectory
                Dim Stringa As String = gf.LeggeFileIntero(Path)
                If Stringa <> "ppp" Then
                    Ritorno = "Errore in fase di lettura disco"
                End If
            End If

            Try
                Kill(Path)
            Catch ex As Exception

            End Try
        End If

        Return Ritorno
    End Function

    Public Function EsegueOperazione(idProc As Integer, Progressivo As Integer, Operazione As Integer, Origine As String, Destinazione As String, Filtro As String,
                            Sovrascrivi As String, SottoDirectory As String, Optional lblOperazione As Label = Nothing,
                            Optional lblContatore As Label = Nothing, Optional UtenzaOrigine As String = "", Optional PasswordOrigine As String = "",
                            Optional UtenzaDest As String = "", Optional PasswordDest As String = "") As StringBuilder
        Dim Gf As New GestioneFilesDirectory
        Dim Filetti() As String = {}
        Dim Cartelle() As String = {}
        Dim qFiletti As Long = -1
        Dim qCartelle As Long = -1
        Dim LeggiCartelle As Boolean = False
        Dim log As StringBuilder = New StringBuilder

        Select Case Operazione
            Case TipoOperazione.Copia
                LeggiCartelle = True
            Case TipoOperazione.Spostamento
                LeggiCartelle = True
            Case TipoOperazione.EliminaDirectory
                LeggiCartelle = True
        End Select

        If Not BloccaTutto Then
            Dim sOrigine As String = ""
            Dim sDestinazione As String = ""
            Dim mud As New MapUnMapDrives
            Dim Oper As New OperazioniSuFileDettagli
            Dim Esegue As Boolean = True
            Dim ope As New OperazioniSuFile

            Dim discoOriginale As String = Mid(Origine, 1, 2)
            Dim PathOriginale As String = Gf.PrendePercorsoDiReteDelDisco(discoOriginale)
            If PathOriginale <> "" Then
                Origine = Origine.Replace(discoOriginale, PathOriginale)
                ope.ScriveLogServizio(idProc, "Rilevato disco di rete di origine:")
                ope.ScriveLogServizio(idProc, "Disco originale: " & discoOriginale)
                ope.ScriveLogServizio(idProc, "Percorso disco:" & PathOriginale)
                ope.ScriveLogServizio(idProc, "Origine: " & Origine)
            End If

            Dim ControlloOrigine As String = ProvaACreareFile(discoOriginale)
            If ControlloOrigine <> "" Then
                ope.ScriveLogServizio(idProc, "Disco di origine " & discoOriginale & " non raggiungibile")
                Esegue = False
            End If

            discoOriginale = Mid(Destinazione, 1, 2)
            PathOriginale = Gf.PrendePercorsoDiReteDelDisco(discoOriginale)
            If PathOriginale <> "" Then
                Destinazione = Destinazione.Replace(discoOriginale, PathOriginale)
                ope.ScriveLogServizio(idProc, "Rilevato disco di rete di destinazione: ")
                ope.ScriveLogServizio(idProc, "Disco originale: " & discoOriginale)
                ope.ScriveLogServizio(idProc, "Percorso disco:" & PathOriginale)
                ope.ScriveLogServizio(idProc, "Destinazione: " & Destinazione)
            End If

            Dim ControlloDestinazione As String = ProvaACreareFile(discoOriginale)
            If ControlloDestinazione <> "" Then
                ope.ScriveLogServizio(idProc, "Disco di destinazione " & discoOriginale & " non raggiungibile")
                Esegue = False
            End If

            If Not BloccaTutto And Esegue Then
                Select Case Operazione
                    Case TipoOperazione.Copia, TipoOperazione.Spostamento
                        log = Oper.Copia(idProc, Operazione, Origine, Destinazione, qFiletti, Filetti, Sovrascrivi, lblOperazione, lblContatore)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.CreaDirectory
                        log = Oper.CreaDirectory(idProc, Origine, Filtro, lblOperazione, lblContatore)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.EliminaDirectory
                        log = Oper.EliminazioneDirectory(idProc, Origine, lblOperazione, lblContatore)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.Eliminazione
                        log = Oper.EliminaFile(idProc, Origine, Filtro, SottoDirectory, lblOperazione, lblContatore)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.SincroniaIntelligente
                        log = Oper.Sincronizza(idProc, Progressivo, DB, ConnSQL, Origine, Destinazione, Filtro, lblOperazione, lblContatore, True)

                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.Sincronizzazione
                        log = Oper.Sincronizza(idProc, Progressivo, DB, ConnSQL, Origine, Destinazione, Filtro, lblOperazione, lblContatore, False)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.RiavvioPC
                        log = Oper.RiavvioPC(idProc)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.FermaServizio
                        FermaServizio(idProc, Origine)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.AvvioServizio
                        FaiPartireServizio(idProc, Origine, Filtro)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.AvviaEseguibile
                        log = Oper.AvviaEseguibile(idProc, Origine, Filtro)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.FermaEseguibile
                        log = Oper.FermaEseguibile(idProc, Origine)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Attendi
                        log = Oper.fAttendi(idProc, Origine)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Zip
                        log = Oper.EsegueZip(idProc, Origine, Destinazione, lblOperazione, lblContatore, True)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.ListaFiles
                        log = Oper.ListaFiles(idProc, Origine, Destinazione, lblOperazione, lblContatore, True)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Messaggio

                        ' ----------------------------------------------------------------------------------------------
                End Select
            End If

            If lblOperazione Is Nothing = False Then
                lblOperazione.Text = ""
            End If

            If lblContatore Is Nothing = False Then
                lblContatore.Text = ""
            End If
            Application.DoEvents()

            mud = Nothing
            Oper = Nothing
            Gf = Nothing
        End If

        Return log
    End Function

    'Public Sub PulisceNomeFileLog()
    '    NomeFileLog = ""
    'End Sub

    'Private Function CreaNomeFileLog() As String
    '    Dim Datella As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day & "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
    '    NomeFileLog = "C:\BackupLog\DettaglioProcedura_" & Datella & ".txt"

    '    Return NomeFileLog
    'End Function

    'Public Function RitornaNomeFileDiLog() As String
    '    If NomeFileLog = "" Then
    '        NomeFileLog = CreaNomeFileLog()
    '    End If
    '    Return NomeFileLog
    'End Function

End Class
