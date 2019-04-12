Imports System.Globalization

Public Class RoutineVarie

    Public Sub ScriveLogServizio(idProc As Integer, Log As String)
        Try
            Dim gf As New GestioneFilesDirectory
            Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
            gf.ApreFileDiTestoPerScrittura("C:\BackupLog\Servizio.txt")
            gf.ScriveTestoSuFileAperto("Procedura: " & idProc & " " & Datella & " -> " & Log)
            gf.ChiudeFileDiTestoDopoScrittura()
            gf = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Function MetteMaiuscoleDopoPunto(Cosa As String) As String
        Dim Ritorno As String = Cosa.ToLower.Trim

        If Ritorno <> "" Then
            If Asc(Mid(Ritorno, 1, 1)) >= Asc("a") And Asc(Mid(Ritorno, 1, 1)) <= Asc("z") Then
                Ritorno = Chr(Asc(Mid(Ritorno, 1, 1)) - 32) & Mid(Ritorno, 2, Len(Ritorno))
            End If
            Ritorno = Ritorno.Replace("  ", " ")
            Ritorno = Ritorno.Replace(". ", ".")
            For i As Integer = 2 To Len(Ritorno)
                If Mid(Ritorno, i, 1) = "." Then
                    If i + 1 < Ritorno.Length Then
                        If Asc(Mid(Ritorno, i + 1, 1)) >= Asc("a") And Asc(Mid(Ritorno, i + 1, 1)) <= Asc("z") Then
                            Ritorno = Mid(Ritorno, 1, i) & Chr(Asc(Mid(Ritorno, i + 1, 1)) - 32) & Mid(Ritorno, i + 2, Len(Ritorno))
                        End If
                    End If
                End If
            Next
            Ritorno = Ritorno.Replace(".", ". ")
        End If

        Return Ritorno
    End Function

    Public Function MetteMaiuscole(Cosa As String) As String
        Dim Ritorno As String = Cosa.ToLower.Trim

        If Ritorno <> "" Then
            If Asc(Mid(Ritorno, 1, 1)) >= Asc("a") And Asc(Mid(Ritorno, 1, 1)) <= Asc("z") Then
                Ritorno = Chr(Asc(Mid(Ritorno, 1, 1)) - 32) & Mid(Ritorno, 2, Len(Ritorno))
            End If
            Ritorno = Ritorno.Replace("  ", " ")
            For i As Integer = 2 To Len(Ritorno)
                If Mid(Ritorno, i, 1) = " " Then
                    If Asc(Mid(Ritorno, i + 1, 1)) >= Asc("a") And Asc(Mid(Ritorno, i + 1, 1)) <= Asc("z") Then
                        Ritorno = Mid(Ritorno, 1, i) & Chr(Asc(Mid(Ritorno, i + 1, 1)) - 32) & Mid(Ritorno, i + 2, Len(Ritorno))
                    End If
                End If
            Next
        End If

        Return Ritorno
    End Function

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

    Public Function Cripta(ByVal Stringa As String) As String
        Dim PassMDB As String = Trim(Stringa)
        Dim VecchiaPass As String = ""

        Randomize()
        Dim Chiave As Integer = Int(Rnd(1) * 64)
        VecchiaPass = Chr(Chiave)
        For i = 1 To Len(PassMDB)
            VecchiaPass = VecchiaPass & Chr(Asc(Mid(PassMDB, i, 1)) + Chiave)
        Next

        Return VecchiaPass
    End Function

    Public Function Decripta(ByVal Stringa As String) As String
        Dim Appoggio As String = ""
        Dim Chiave As Integer

        Try
            Chiave = Asc(Left(Trim(Stringa), 1))
            For i = 2 To Len(Stringa)
                Appoggio = Appoggio & Chr(Asc(Mid(Stringa, i, 1)) - Chiave)
            Next
            Appoggio = Trim(Appoggio)
        Catch ex As Exception
            Appoggio = ""
        End Try

        Return Appoggio
    End Function

    Public Function SistemaTestoPerDB(Cosa As String) As String
        Dim Ritorno As String = "" & Cosa

        Ritorno = Ritorno.Replace("'", "''")

        Return Ritorno
    End Function

    Public Function ConverteDataPerDB(Datella As String) As String
        If Datella = "" Then
            Return "null"
        End If

        Dim Datellaccia As Date = Datella

        If (DBNull.Value.Equals(Datella)) = True Then
            Return "null"
        Else
            Return "'" & Datellaccia.Year & "-" & Format(Datellaccia.Month, "00") & "-" & Format(Datellaccia.Day, "00") & " " & Format(Datellaccia.Hour, "00") & ":" & Format(Datellaccia.Minute, "00") & ":" & Format(Datellaccia.Second, "00") & ".000" & "'"
        End If
    End Function

    Public Function ConverteEsadecimaleADecimale(NumeroHex As String) As Integer
        Dim i As Integer = Convert.ToInt32(NumeroHex, 16)

        Return i
    End Function

    Public Function ConvertePercentualiSuIndirizzo(Cosa As String) As String
        Dim Appo As String = Cosa
        Dim Numerello As String
        Dim C As Char
        Dim N As Integer
        Dim N2 As String
        Dim Dove As Long

        Dove = Appo.IndexOf("%")
        Do While Dove > -1
            Numerello = Mid(Appo, Dove + 1, 3)
            Try
                N2 = Mid(Numerello, 2, 2)
                N = Convert.ToInt32(N2, 16)
            Catch ex As Exception
                N = -1
            End Try
            If N <> -1 Then
                C = Chr(N)
            End If
            Appo = Appo.Replace(Numerello, C)

            Dove = Appo.IndexOf("%")
        Loop

        Appo = Appo.Replace("&quot;", Chr(34))
        Appo = Appo.Replace("$", "_")
        Appo = Appo.Replace("\/", "/")
        Appo = Appo.Replace("\\/", "/")

        Appo = Appo.Replace("&amp;", "&")
        Appo = Appo.Replace("&gt;", "<")
        Appo = Appo.Replace("&lt;", ">")
        Appo = Appo.Replace("&quot;", Chr(34))
        Appo = Appo.Replace("&tilde;", "˜")

        Return Appo
    End Function
End Class
