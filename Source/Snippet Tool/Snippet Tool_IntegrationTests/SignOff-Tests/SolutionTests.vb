
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
Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Microsoft.VSSDK.Tools.VsIdeTesting
Imports EnvDTE

Namespace Snippet_Tool_IntegrationTests.IntegrationTests

    <TestClass()> _
    Public Class SolutionTests

        Private Delegate Sub ThreadInvoker()
        Private _testContext As TestContext

        ''' <summary>
        '''Gets or sets the test context which provides
        '''information about and functionality for the current test run.
        '''</summary>
        Public Property TestContext() As TestContext
            Get
                Return _testContext
            End Get
            Set(ByVal value As TestContext)
                _testContext = value
            End Set
        End Property

        <TestMethod(), HostType("VS IDE")> _
        Public Sub CreateEmptySolution()
            UIThreadInvoker.Invoke(CType(AddressOf AnonymousMethod1, ThreadInvoker))
        End Sub
        Private Sub AnonymousMethod1()
            Dim testUtils As Microsoft.VsSDK.IntegrationTestLibrary.TestUtils = New Microsoft.VsSDK.IntegrationTestLibrary.TestUtils()
            testUtils.CloseCurrentSolution(__VSSLNSAVEOPTIONS.SLNSAVEOPT_NoSave)
            testUtils.CreateEmptySolution(TestContext.TestDir, "EmptySolution")
        End Sub

    End Class
End Namespace
