
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
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Microsoft.VSSDK.Tools.VsIdeTesting
Imports EnvDTE

Namespace Snippet_Tool_IntegrationTests.IntegrationTests

    <TestClass()> _
    Public Class CppProjectTests
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
        Public Sub CPPWinformsApplication()
            UIThreadInvoker.Invoke(CType(AddressOf AnonymousMethod1, ThreadInvoker))
        End Sub
        Private Sub AnonymousMethod1()

            'Solution and project creation parameters
            Dim solutionName As String = "CPPWinApp"
            Dim projectName As String = "CPPWinApp"

            'Template parameters
            Dim projectType As String = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}"
            Dim projectTemplateName As String = Path.Combine("vcNet", "mc++appwiz.vsz")
            Dim itemTemplateName As String = "newc++file.cpp"
            Dim newFileName As String = "Test.cpp"


            Dim dte As DTE = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(DTE)), DTE)

            Dim testUtils As Microsoft.VsSDK.IntegrationTestLibrary.TestUtils = New Microsoft.VsSDK.IntegrationTestLibrary.TestUtils()

            testUtils.CreateEmptySolution(TestContext.TestDir, solutionName)
            Assert.AreEqual(Of Integer)(0, testUtils.ProjectCount())

            'Add new CPP Windows application project to existing solution
            Dim solutionDirectory As String = Directory.GetParent(dte.Solution.FullName).FullName
            Dim projectDirectory As String = Microsoft.VsSDK.IntegrationTestLibrary.TestUtils.GetNewDirectoryName(solutionDirectory, projectName)
            Dim projectTemplatePath As String = Path.Combine(dte.Solution.TemplatePath(projectType), projectTemplateName)
            Assert.IsTrue(File.Exists(projectTemplatePath), String.Format("Could not find template file: {0}", projectTemplatePath))
            dte.Solution.AddFromTemplate(projectTemplatePath, projectDirectory, projectName, False)

            'Verify that the new project has been added to the solution
            Assert.AreEqual(Of Integer)(1, testUtils.ProjectCount())

            'Get the project
            Dim project As Project = dte.Solution.Item(1)
            Assert.IsNotNull(project)
            Assert.IsTrue(String.Compare(project.Name, projectName, StringComparison.InvariantCultureIgnoreCase) = 0)

            'Verify Adding new code file to project
            Dim newItemTemplatePath As String = Path.Combine(dte.Solution.ProjectItemsTemplatePath(projectType), itemTemplateName)
            Assert.IsTrue(File.Exists(newItemTemplatePath))
            Dim item As ProjectItem = project.ProjectItems.AddFromTemplate(newItemTemplatePath, newFileName)
            Assert.IsNotNull(item)

        End Sub

    End Class
End Namespace
