Imports System.ComponentModel
Imports System.Configuration.Install

Public Class ProjectInstaller

    Public Sub New()
        MyBase.New()

        'Chiamata richiesta da Progettazione componenti.
        InitializeComponent()

        'Aggiungere il codice di inizializzazione dopo la chiamata a InitializeComponent
    End Sub

    Public Overrides Sub Commit(ByVal savedState As System.Collections.IDictionary)
        'Dim ckey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
        'ckey.OpenSubKey("SYSTEM\CurrentControlSet\Services\ServiceDemoCalc", True)

        'If ckey Is Nothing Then
        '    MsgBox("La chiave non esiste errore di sistema", MsgBoxStyle.Information)
        'Else
        '    If DirectCast(ckey.GetValueNames, IList).IndexOf("Type") = -1 Then
        '        ckey.OpenSubKey("SYSTEM\CurrentControlSet\Services\ServiceDemoCalc", True).SetValue("Type", CType(272, Integer))
        '    Else
        '        ckey.OpenSubKey("SYSTEM\CurrentControlSet\Services\SeerviceDemoCalc", True).SetValue("Type", CType(272, Integer))
        '        MsgBox("Registrato Servizio Correttamente", MsgBoxStyle.Information)
        '    End If
        'End If
    End Sub

    Private Sub ServiceProcessInstaller1_AfterInstall(sender As Object, e As InstallEventArgs) Handles ServiceProcessInstaller1.AfterInstall

    End Sub
End Class
