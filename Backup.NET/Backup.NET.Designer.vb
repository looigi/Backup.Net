Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BackupNET
    Inherits System.ServiceProcess.ServiceBase

    'UserService esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Il punto di ingresso principale del processo
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' All'interno di uno stesso processo è possibile eseguire più servizi di Windows NT.
        ' Per aggiungere un servizio al processo, modificare la riga che segue in modo
        ' da creare un secondo oggetto servizio. Ad esempio,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New BackupNET}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Richiesto da Progettazione componenti
    Private components As System.ComponentModel.IContainer

    ' NOTA: la procedura che segue è richiesta da Progettazione componenti
    ' Può essere modificata in Progettazione componenti.  
    ' Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        '
        'Timer2
        '
        '
        'BackupNET
        '
        Me.ServiceName = "Backup.NET"

    End Sub
    Friend WithEvents Timer2 As System.Windows.Forms.Timer

End Class
