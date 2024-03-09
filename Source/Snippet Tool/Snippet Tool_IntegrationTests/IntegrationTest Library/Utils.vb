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
Imports System.Reflection
Imports System.Diagnostics
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.Win32
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Microsoft.VSSDK.Tools.VsIdeTesting
Imports Microsoft.VisualStudio
Imports EnvDTE
Imports EnvDTE80

Namespace Microsoft.VsSDK.IntegrationTestLibrary

    Public Class TestUtils

#Region "Methods: Handling embedded resources"
        ''' <summary>
        ''' Gets the embedded file identified by the resource name, and converts the
        ''' file into a string.
        ''' </summary>
        ''' <param name="resourceName">In VS, is DefaultNamespace.FileName?</param>
        ''' <returns></returns>
        Public Shared Function GetEmbeddedStringResource(ByVal [assembly] As System.Reflection.Assembly, ByVal resourceName As String) As String
            Dim result As String = Nothing

            ' Use the .NET procedure for loading a file embedded in the assembly
            Dim stream As Stream = [assembly].GetManifestResourceStream(resourceName)
            If Not stream Is Nothing Then
                ' Convert bytes to string
                Dim fileContentsAsBytes As Byte() = New Byte(CInt(stream.Length - 1)) {}
                stream.Read(fileContentsAsBytes, 0, CInt(Fix(stream.Length)))
                result = Encoding.Default.GetString(fileContentsAsBytes)
            Else
                ' Embedded resource not found - list available resources
                Debug.WriteLine("Unable to find the embedded resource file '" & resourceName & "'.")
                Debug.WriteLine("  Available resources:")
                For Each aResourceName As String In [assembly].GetManifestResourceNames()
                    Debug.WriteLine("    " & aResourceName)
                Next aResourceName
            End If

            Return result
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="embeddedResourceName"></param>
        ''' <param name="baseFileName"></param>
        ''' <param name="fileExtension"></param>
        ''' <returns></returns>
        Public Shared Sub WriteEmbeddedResourceToFile(ByVal [assembly] As System.Reflection.Assembly, ByVal embeddedResourceName As String, ByVal fileName As String)
            ' Get file contents
            Dim fileContents As String = GetEmbeddedStringResource([assembly], embeddedResourceName)
            If fileContents Is Nothing Then
                Throw New ApplicationException("Failed to get embedded resource '" & embeddedResourceName & "' from assembly '" & [assembly].FullName)
            End If

            ' Write to file
            Dim sw As StreamWriter = New StreamWriter(fileName)
            sw.Write(fileContents)
            sw.Close()
        End Sub

        ''' <summary>
        ''' Writes an embedded resource to a file.
        ''' </summary>
        ''' <param name="assembly">The name of the assembly that the embedded resource is defined.</param>
        ''' <param name="embeddedResourceName">The name of the embedded resource.</param>
        ''' <param name="fileName">The file to write the embedded resource's content.</param>
        Public Shared Sub WriteEmbeddedResourceToBinaryFile(ByVal [assembly] As System.Reflection.Assembly, ByVal embeddedResourceName As String, ByVal fileName As String)
            ' Get file contents
            Dim stream As Stream = [assembly].GetManifestResourceStream(embeddedResourceName)
            If stream Is Nothing Then
                Throw New InvalidOperationException("Failed to get embedded resource '" & embeddedResourceName & "' from assembly '" & [assembly].FullName)
            End If

            ' Write to file
            Dim sw As BinaryWriter = Nothing
            Dim fs As FileStream = Nothing
            Try
                Dim fileContentsAsBytes As Byte() = New Byte(CInt(stream.Length - 1)) {}
                stream.Read(fileContentsAsBytes, 0, CInt(Fix(stream.Length)))

                Dim mode As FileMode = FileMode.CreateNew
                If File.Exists(fileName) Then
                    mode = FileMode.Truncate
                End If

                fs = New FileStream(fileName, mode)

                sw = New BinaryWriter(fs)
                sw.Write(fileContentsAsBytes)
            Finally
                If Not fs Is Nothing Then
                    fs.Close()
                End If
                If Not sw Is Nothing Then
                    sw.Close()
                End If
            End Try
        End Sub

#End Region

#Region "Methods: Handling temporary files and directories"
        ''' <summary>
        ''' Returns the first available file name on the form
        '''   [baseFileName]i.[extension]
        ''' where [i] starts at 1 and increases until there is an available file name
        ''' in the given directory. Also creates an empty file with that name to mark
        ''' that file as occupied.
        ''' </summary>
        ''' <param name="directory">Directory that the file should live in.</param>
        ''' <param name="baseFileName"></param>
        ''' <param name="extension">may be null, in which case the .[extension] part
        ''' is not added.</param>
        ''' <returns>Full file name.</returns>
        Public Shared Function GetNewFileName(ByVal directory As String, ByVal baseFileName As String, ByVal extension As String) As String
            ' Get the new file name
            Dim fileName As String = GetNewFileOrDirectoryNameWithoutCreatingAnything(directory, baseFileName, extension)

            ' Create an empty file to mark it as taken
            Dim sw As StreamWriter = New StreamWriter(fileName)

            sw.Write("")
            sw.Close()
            Return fileName
        End Function
        ''' <summary>
        ''' Returns the first available directory name on the form
        '''   [baseDirectoryName]i
        ''' where [i] starts at 1 and increases until there is an available directory name
        ''' in the given directory. Also creates the directory to mark it as occupied.
        ''' </summary>
        ''' <param name="directory">Directory that the file should live in.</param>
        ''' <param name="baseDirectoryName"></param>
        ''' <returns>Full directory name.</returns>
        Public Shared Function GetNewDirectoryName(ByVal directory As String, ByVal baseDirectoryName As String) As String
            ' Get the new file name
            Dim directoryName As String = GetNewFileOrDirectoryNameWithoutCreatingAnything(directory, baseDirectoryName, Nothing)

            ' Create an empty directory to make it as occupied
            System.IO.Directory.CreateDirectory(directoryName)

            Return directoryName
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="directory"></param>
        ''' <param name="baseFileName"></param>
        ''' <param name="extension"></param>
        ''' <returns></returns>
        Private Shared Function GetNewFileOrDirectoryNameWithoutCreatingAnything(ByVal directory As String, ByVal baseFileName As String, ByVal extension As String) As String
            ' - get a file name that we can use
            Dim fileName As String
            Dim i As Integer = 1

            Dim fullFileName As String = Nothing
            Do
                ' construct next file name
                fileName = baseFileName & i
                If Not extension Is Nothing Then
                    fileName &= "."c & extension
                End If

                ' check if that file exists in the directory
                fullFileName = Path.Combine(directory, fileName)

                If (Not File.Exists(fullFileName)) AndAlso (Not System.IO.Directory.Exists(fullFileName)) Then
                    Exit Do
                Else
                    i += 1
                End If
            Loop

            Return fullFileName
        End Function
#End Region

#Region "Methods: Handling solutions"
        ''' <summary>
        ''' Closes the currently open solution (if any), and creates a new solution with the given name.
        ''' </summary>
        ''' <param name="solutionName">Name of new solution.</param>
        Public Sub CreateEmptySolution(ByVal directory As String, ByVal solutionName As String)
            CloseCurrentSolution(__VSSLNSAVEOPTIONS.SLNSAVEOPT_NoSave)

            Dim solutionDirectory As String = GetNewDirectoryName(directory, solutionName)

            ' Create and force save solution
            Dim solutionService As IVsSolution = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)
            solutionService.CreateSolution(solutionDirectory, solutionName, CUInt(__VSCREATESOLUTIONFLAGS.CSF_SILENT))
            solutionService.SaveSolutionElement(CUInt(__VSSLNSAVEOPTIONS.SLNSAVEOPT_ForceSave), Nothing, 0)
            Dim dte As DTE = VsIdeTestHostContext.Dte
            Assert.AreEqual(solutionName & ".sln", Path.GetFileName(dte.Solution.FileName), "Newly created solution has wrong Filename")
        End Sub

        Public Sub CloseCurrentSolution(ByVal saveoptions As __VSSLNSAVEOPTIONS)
            ' Get solution service
            Dim solutionService As IVsSolution = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)

            ' Close already open solution
            solutionService.CloseSolutionElement(CUInt(saveoptions), Nothing, 0)
        End Sub

        Public Sub ForceSaveSolution()
            ' Get solution service
            Dim solutionService As IVsSolution = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)

            ' Force-save the solution
            solutionService.SaveSolutionElement(CUInt(__VSSLNSAVEOPTIONS.SLNSAVEOPT_ForceSave), Nothing, 0)
        End Sub

        ''' <summary>
        ''' Get current number of open project in solution
        ''' </summary>
        ''' <returns></returns>
        Public Function ProjectCount() As Integer
            ' Get solution service
            Dim solutionService As IVsSolution = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)
            Dim projectCount_Renamed As Object = Nothing
            solutionService.GetProperty(CInt(Fix(__VSPROPID.VSPROPID_ProjectCount)), projectCount_Renamed)
            Return CInt(Fix(projectCount_Renamed))
        End Function
#End Region

#Region "Methods: Handling projects"
        ''' <summary>
        ''' Creates a project.
        ''' </summary>
        ''' <param name="projectName">Name of new project.</param>
        ''' <param name="templateName">Name of project template to use</param>
        ''' <param name="language">language</param>
        ''' <returns>New project.</returns>
        Public Sub CreateProjectFromTemplate(ByVal projectName As String, ByVal templateName As String, ByVal language As String, ByVal exclusive As Boolean)
            Dim dte As DTE = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(DTE)), DTE)

            Dim sol As Solution2 = TryCast(dte.Solution, Solution2)
            Dim projectTemplate As String = sol.GetProjectTemplate(templateName, language)

            ' - project name and directory
            Dim solutionDirectory As String = Directory.GetParent(dte.Solution.FullName).FullName
            Dim projectDirectory As String = GetNewDirectoryName(solutionDirectory, projectName)

            dte.Solution.AddFromTemplate(projectTemplate, projectDirectory, projectName, False)
        End Sub
#End Region

#Region "Methods: Handling project items"
        ''' <summary>
        ''' Create a new item in the project
        ''' </summary>
        ''' <param name="parent">the parent collection for the new item</param>
        ''' <param name="templateName"></param>
        ''' <param name="language"></param>
        ''' <param name="name"></param>
        ''' <returns></returns>
        Public Function AddNewItemFromVsTemplate(ByVal parent As ProjectItems, ByVal templateName As String, ByVal language As String, ByVal name As String) As ProjectItem
            If parent Is Nothing Then
                Throw New ArgumentException("project")
            End If
            If name Is Nothing Then
                Throw New ArgumentException("name")
            End If

            Dim dte As DTE = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(DTE)), DTE)

            Dim sol As Solution2 = TryCast(dte.Solution, Solution2)

            Dim filename As String = sol.GetProjectItemTemplate(templateName, language)

            parent.AddFromTemplate(filename, name)

            Return parent.Item(name)
        End Function

        ''' <summary>
        ''' Save an open document.
        ''' </summary>
        ''' <param name="documentMoniker">for filebased documents this is the full path to the document</param>
        Public Sub SaveDocument(ByVal documentMoniker As String)
            ' Get document cookie and hierarchy for the file
            Dim runningDocumentTableService As IVsRunningDocumentTable = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsRunningDocumentTable)), IVsRunningDocumentTable)
            Dim docCookie As UInteger
            Dim docData As IntPtr
            Dim hierarchy As IVsHierarchy = Nothing
            Dim itemId As UInteger
            runningDocumentTableService.FindAndLockDocument(CUInt(_VSRDTFLAGS.RDT_NoLock), documentMoniker, hierarchy, itemId, docData, docCookie)

            ' Save the document
            Dim solutionService As IVsSolution = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)
            solutionService.SaveSolutionElement(CUInt(__VSSLNSAVEOPTIONS.SLNSAVEOPT_ForceSave), hierarchy, docCookie)
        End Sub

        Public Sub CloseInEditorWithoutSaving(ByVal fullFileName As String)
            ' Get the RDT service
            Dim runningDocumentTableService As IVsRunningDocumentTable = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsRunningDocumentTable)), IVsRunningDocumentTable)
            Assert.IsNotNull(runningDocumentTableService, "Failed to get the Running Document Table Service")

            ' Get our document cookie and hierarchy for the file
            Dim docCookie As UInteger
            Dim docData As IntPtr
            Dim hierarchy As IVsHierarchy = Nothing
            Dim itemId As UInteger
            runningDocumentTableService.FindAndLockDocument(CUInt(_VSRDTFLAGS.RDT_NoLock), fullFileName, hierarchy, itemId, docData, docCookie)

            ' Get the SolutionService
            Dim solutionService As IVsSolution = TryCast(VsIdeTestHostContext.ServiceProvider.GetService(GetType(IVsSolution)), IVsSolution)
            Assert.IsNotNull(solutionService, "Failed to get IVsSolution service")

            ' Close the document
            solutionService.CloseSolutionElement(CUInt(__VSSLNSAVEOPTIONS.SLNSAVEOPT_NoSave), hierarchy, docCookie)
        End Sub
#End Region

#Region "Methods: Handling Toolwindows"
        Public Function CanFindToolwindow(ByVal persistenceGuid As Guid) As Boolean
            Dim uiShellService As IVsUIShell = TryCast(VsIdeTestHostContext.ServiceProvider.GetService(GetType(SVsUIShell)), IVsUIShell)
            Assert.IsNotNull(uiShellService)
            Dim windowFrame As IVsWindowFrame = Nothing
            Dim hr As Integer = uiShellService.FindToolWindow(CUInt(__VSFINDTOOLWIN.FTW_fFindFirst), persistenceGuid, windowFrame)
            Assert.IsTrue(hr = VSConstants.S_OK)

            Return (Not windowFrame Is Nothing)
        End Function
#End Region

#Region "Methods: Loading packages"
        Public Function LoadPackage(ByVal packageGuid As Guid) As IVsPackage
            Dim shellService As IVsShell = CType(VsIdeTestHostContext.ServiceProvider.GetService(GetType(SVsShell)), IVsShell)
            Dim package As IVsPackage = Nothing
            shellService.LoadPackage(packageGuid, package)
            Assert.IsNotNull(package, "Failed to load package")
            Return package
        End Function
#End Region

        ''' <summary>
        ''' Executes a Command (menu item) in the given context
        ''' </summary>
        Public Sub ExecuteCommand(ByVal cmd As CommandID)
            Dim Customin As Object = Nothing
            Dim Customout As Object = Nothing
            Dim guidString As String = cmd.Guid.ToString("B").ToUpper()
            Dim cmdId As Integer = cmd.ID
            Dim dte As DTE = VsIdeTestHostContext.Dte
            dte.Commands.Raise(guidString, cmdId, Customin, Customout)
        End Sub

    End Class
End Namespace
