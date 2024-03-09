'**************************************************************************

'Copyright (c) Microsoft Corporation. All rights reserved.
'This code is licensed under the Visual Studio SDK license terms.
'THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
'ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
'IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
'PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

'**************************************************************************


Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VSSDK.Tools.VsIdeTesting
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports EnvDTE
Imports ElektroStudios.SnippetToolPackage

Namespace Snippet_Tool_IntegrationTests
    ''' <summary>
    ''' Summary description for PackageTest
    ''' </summary>
    <TestClass()> _
    Public Class PackageTest

        Private Delegate Sub ThreadInvoker()

        <TestMethod(), HostType("VS IDE")> _
        Public Sub PackageLoadTest()
            Dim objThreadInvoker As New ThreadInvoker(AddressOf AnonymousMethod1)
            UIThreadInvoker.Invoke(objThreadInvoker)
        End Sub

        ''' <summary>
        ''' </summary>
        Private Sub AnonymousMethod1()

            ' Get the shell service
            Dim sp As IServiceProvider = VsIdeTestHostContext.ServiceProvider
            Dim shellService As IVsShell = TryCast(sp.GetService(GetType(SVsShell)), IVsShell)
            Assert.IsNotNull(shellService)

            'Validate package load
            Dim package As IVsPackage = Nothing
            Dim packageGuid As Guid = New Guid(MyPackage.Info.GuidPackageString)
            Assert.IsTrue(0 = shellService.LoadPackage(packageGuid, package))

        End Sub

    End Class
End Namespace
