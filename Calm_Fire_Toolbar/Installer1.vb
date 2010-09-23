Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices


Partial Public Class Installer1
    Inherits Installer

    Public Sub New()
        InitializeComponent()
    End Sub


    Public Overloads Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
        Dim regSrv As New RegistrationServices()
        regSrv.RegisterAssembly(MyBase.[GetType]().Assembly, AssemblyRegistrationFlags.SetCodeBase)
    End Sub

    Public Overloads Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        MyBase.Uninstall(savedState)
        Dim regSrv As New RegistrationServices()
        regSrv.UnregisterAssembly(MyBase.[GetType]().Assembly)
    End Sub


End Class
