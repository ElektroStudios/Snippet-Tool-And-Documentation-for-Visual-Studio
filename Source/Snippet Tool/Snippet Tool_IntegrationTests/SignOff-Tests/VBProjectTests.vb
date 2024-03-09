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
Imports EnvDTE

Namespace Snippet_Tool_IntegrationTests.IntegrationTests

    <TestClass()> _
    Public Class VisualBasicProjectTests
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
        Public Sub VBWinformsApplication()
            UIThreadInvoker.Invoke(CType(AddressOf AnonymousMethod1, ThreadInvoker))
        End Sub

        Private Sub AnonymousMethod1()

            'Solution and project creation parameters
            Dim solutionName As String = "VBWinApp"
            Dim projectName As String = "VBWinApp"

            'Template parameters
            Dim language As String = "VisualBasic"
            Dim projectTemplateName As String = "WindowsApplication.Zip"
            Dim itemTemplateName As String = "CodeFile.zip"
            Dim newFileName As String = "Test.vb"

            Dim dte As DTE = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(DTE)), DTE)

            Dim testUtils As Microsoft.VsSDK.IntegrationTestLibrary.TestUtils = New Microsoft.VsSDK.IntegrationTestLibrary.TestUtils()

            testUtils.CreateEmptySolution(TestContext.TestDir, solutionName)
            Assert.AreEqual(Of Integer)(0, testUtils.ProjectCount())

            'Add new  Windows application project to existing solution
            testUtils.CreateProjectFromTemplate(projectName, projectTemplateName, language, False)

            'Verify that the new project has been added to the solution
            Assert.AreEqual(Of Integer)(1, testUtils.ProjectCount())

            'Get the project
            Dim project As Project = dte.Solution.Item(1)
            Assert.IsNotNull(project)
            Assert.IsTrue(String.Compare(project.Name, projectName, StringComparison.InvariantCultureIgnoreCase) = 0)

            'Verify Adding new code file to project
            Dim newCodeFileItem As ProjectItem = testUtils.AddNewItemFromVsTemplate(project.ProjectItems, itemTemplateName, language, newFileName)
            Assert.IsNotNull(newCodeFileItem, "Could not create new project item")
        End Sub

    End Class
End Namespace
