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
Imports System.Text
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Microsoft.VSSDK.Tools.VsIdeTesting

Namespace Snippet_Tool_IntegrationTests.IntegrationTests

    <TestClass()> _
    Public Class CSharpProjectTests
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
        Public Sub WinformsApplication()
            UIThreadInvoker.Invoke(CType(AddressOf AnonymousMethod1, ThreadInvoker))
        End Sub
        Private Sub AnonymousMethod1()
            Dim testUtils As Microsoft.VsSDK.IntegrationTestLibrary.TestUtils = New Microsoft.VsSDK.IntegrationTestLibrary.TestUtils()
            testUtils.CreateEmptySolution(TestContext.TestDir, "CSWinApp")
            Assert.AreEqual(Of Integer)(0, testUtils.ProjectCount())
        End Sub

    End Class
End Namespace
