Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.ServiceProcess

<RunInstaller(True)>
Public Class ProjectInstaller
    Inherits Installer

    Private serviceInstaller As ServiceInstaller
    Private processInstaller As ServiceProcessInstaller

    Public Sub New()
        processInstaller = New ServiceProcessInstaller()
        processInstaller.Account = ServiceAccount.LocalSystem

        serviceInstaller = New ServiceInstaller()
        serviceInstaller.ServiceName = "PAS-Service"
        serviceInstaller.DisplayName = "PAS Web Service"
        serviceInstaller.Description = "PAS Medical Office Integration Service"
        serviceInstaller.StartType = ServiceStartMode.Automatic

        Installers.Add(processInstaller)
        Installers.Add(serviceInstaller)
    End Sub
End Class
