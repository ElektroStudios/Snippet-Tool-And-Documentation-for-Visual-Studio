'**************************************************************************
'Copyright (c) Microsoft Corporation. All rights reserved.
'This code is licensed under the Visual Studio SDK license terms.
'THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
'ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
'IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
'PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
'**************************************************************************
Imports System
Imports System.ComponentModel.Design
Imports System.Globalization
Imports Microsoft.VSSDK.Tools.VsIdeTesting
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports ElektroStudios.SnippetToolPackage

''' <summary>
''' Summary description for MenuItemTest
''' </summary>
<TestClass()>
Public Class MenuItemTest

    Private Delegate Sub ThreadInvoker()

    <TestMethod(), HostType("VS IDE")> _
    Public Sub LaunchCommand()
        Dim objThreadInvoker As New ThreadInvoker(AddressOf AnonymousMethod1)
        UIThreadInvoker.Invoke(objThreadInvoker)
    End Sub

    ''' <summary>
    ''' </summary>
    Private Sub AnonymousMethod1()

        Dim menuItemCmd As CommandID = New CommandID(MyPackage.Info.GuidCmdSetRef, CInt(MyPackage.Info.CmdIdCRef))

        ' Create the DialogBoxListener Thread.
        Dim expectedDialogBoxText As String = String.Format(CultureInfo.CurrentUICulture, "{0}" & _
                                                            vbLf & _
                                                            vbLf & _
                                                            "Inside {1}.MenuItemCallback()", "Snippet Tool", "Snippet_ToolPackage")
        Dim purger As IntegrationTest_Library.DialogBoxPurger = New IntegrationTest_Library.DialogBoxPurger(Microsoft.VsSDK.IntegrationTestLibrary.NativeMethods.IDOK, expectedDialogBoxText)

        Try
            purger.Start()

            ExecuteCommand(menuItemCmd)

        Finally
            Assert.IsTrue(purger.WaitForDialogThreadToTerminate(), "The dialog box has not been shown")
        End Try
    End Sub

    Private Sub ExecuteCommand(ByVal cmd As CommandID)
        Dim Customin As Object = Nothing
        Dim Customout As Object = Nothing
        Dim guidString As String = cmd.Guid.ToString("B").ToUpper()
        Dim cmdId As Integer = cmd.ID
        Dim dte As EnvDTE.DTE = VsIdeTestHostContext.Dte
        dte.Commands.Raise(guidString, cmdId, Customin, Customout)
    End Sub

End Class

