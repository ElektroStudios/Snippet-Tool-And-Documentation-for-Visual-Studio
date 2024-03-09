' ***********************************************************************
' Author   : Elektro
' Modified : 11-December-2015
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell

Imports ElektroStudios.SnippetToolPackage.MyPackage.Enums
Imports ElektroStudios.SnippetToolPackage.MyPackage.Tools
Imports ElektroStudios.SnippetToolPackage.MyPackage.UserInterface

#End Region

#Region " Main "

Namespace MyPackage.Package

    <ProvideMenuResource("Menus.ctmenu", 1)>
    <ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)>
    <ProvideOptionPage(GetType(SnippetsPageGrid), "Snippet Tool", "Snippet Defaults", 0, 0, True)>
    <ProvideOptionPage(GetType(XmlDocPageGrid), "Snippet Tool", "Xml Defaults", 0, 0, True)>
    Partial Public NotInheritable Class Main : Inherits AsyncPackage

#Region " CallBacks ( Commands ) "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdCRef"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdCRef_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeCodeRef), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdParamRef"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdParamRef_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeParamRef), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdLangRef"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdLangRef_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeLangRef), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdSinglelineCode"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdSinglelineCode_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeSinglelineCode), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdMultilineCode"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdMultilineCode_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeMultilineCode), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdCodeExample"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdCodeExample_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeCodeExample), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdLink"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdLink_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeLink), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdLinkAlter"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdLinkAlter_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeLinkAlter), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdSeparator"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdSeparator_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeSeparator), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdParagraph"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdParagraph_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeParagraph), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdRemarks"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdRemarks_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeRemarks), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdCollapse"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdCollapse_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.CollapseXmlComments), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdExpand"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdExpand_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.ExpandXmlComments), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdDelete"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdDelete_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.DeleteXmlComments), {0})

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' <see cref="CmdCreateSnippet"/> callback handler.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="sender">
        ''' The source of the event.
        ''' </param>
        ''' 
        ''' <param name="e">
        ''' The <see cref="EventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub CmdCreateSnippet_Callback(sender As Object, e As EventArgs)

            ErrorHandler.ThrowOnFailure(Documentation.ProcessSelectedText(Main.Dte, CodeEditor.GetCurrentViewHost, DocumentationTask.MakeSnippet), {0})

        End Sub

#End Region

    End Class

End Namespace

#End Region
