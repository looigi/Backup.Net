Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime

Public Class GestioneMAIL
    Public Sub InvioMAIL(opFiles As OperazioniSuFile, idProc As Integer, Oggetto As String, Testo As String, Optional sAttachment As String = "")
        Dim errore As Boolean = False
        Dim Utenza As String = LeggeUtenzaPassword()
        If Utenza <> "" Then
            Dim campiUtenza() As String = Utenza.Split(";")

            opFiles.ScriveLogServizio(idProc, "Invio mail: destinatari: " & Utenza)

            Try
                Dim Mail As New MailMessage

                Mail.Subject = Oggetto
                Mail.From = New MailAddress(campiUtenza(0))
                Mail.To.Add(campiUtenza(0))
                Mail.Body = Testo 'Message Here
                If sAttachment <> "" Then
                    If File.Exists(sAttachment) Then
                        Try
                            Dim data As Attachment = New Attachment(sAttachment, MediaTypeNames.Application.Octet)

                            Mail.Attachments.Add(data)
                        Catch ex As Exception
                            opFiles.ScriveLogServizio(idProc, "ERRORE: Problemi nell'inserire nella mail il file " & sAttachment & ": " & ex.Message)
                        End Try
                    End If
                End If

                Dim SMTP As New SmtpClient("smtp.gmail.com")

                SMTP.Credentials = New System.Net.NetworkCredential(campiUtenza(0), campiUtenza(1))
                SMTP.Port = "587"
                SMTP.EnableSsl = True

                opFiles.ScriveLogServizio(idProc, "Invio mail in corso")

                SMTP.Send(Mail)
            Catch ex As SmtpException
                opFiles.ScriveLogServizio(idProc, "ERRORE Invio mail: " & ex.Message)
                errore = True
            Catch ex As Exception
                opFiles.ScriveLogServizio(idProc, "ERRORE Invio mail: " & ex.Message)
                errore = True
            End Try

            If Not errore Then
                If sAttachment <> "" Then
                    Try
                        Kill(sAttachment)
                    Catch ex As Exception

                    End Try
                End If
            End If
            opFiles.ScriveLogServizio(idProc, "Invio mail effettuato")
        End If
    End Sub
End Class
