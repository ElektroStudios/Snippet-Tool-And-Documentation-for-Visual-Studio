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

Imports ElektroStudios.SnippetToolPackage.MyPackage.Tools
Imports ElektroStudios.SnippetToolPackage.MyPackage.UserInterface

#End Region

Namespace MyPackage.Package

    <ProvideMenuResource("Menus.ctmenu", 1)>
    <ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)>
    <ProvideOptionPage(GetType(SnippetsPageGrid), "Snippet Tool", "Snippet Defaults", 0, 0, True)>
    <ProvideOptionPage(GetType(XmlDocPageGrid), "Snippet Tool", "Xml Defaults", 0, 0, True)>
    Partial Public NotInheritable Class Main : Inherits AsyncPackage

#Region " Event-Handlers ( Commands ) "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdCRef"/> command.
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
        Private Sub CmdCRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdCRef.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled = CodeEditor.IsCaretOnXmlBlock()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdParamRef"/> command.
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
        Private Sub CmdParamRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdParamRef.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled = CodeEditor.IsCaretOnXmlBlock()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdLangRef"/> command.
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
        Private Sub CmdLangRef_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdLangRef.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled = CodeEditor.IsCaretOnXmlBlock()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdSinglelineCode"/> command.
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
        Private Sub CmdSinglelineCode_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdSinglelineCode.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                CodeEditor.IsCaretOnXmlBlock AndAlso
                CodeEditor.IsTextSelectedSingleLine()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdMultilineCode"/> command.
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
        Private Sub CmdMultilineCode_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdMultilineCode.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                CodeEditor.IsCaretOnXmlBlock AndAlso
                CodeEditor.IsTextSelectedMultiLine()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdCodeExample"/> command.
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
        Private Sub CmdCodeExample_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdCodeExample.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled = CodeEditor.IsTextSelected()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdLink"/> command.
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
        Private Sub CmdLink_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdLink.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
               CodeEditor.IsCaretOnXmlBlock AndAlso
                CodeEditor.IsTextSelectedSingleLine()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdLinkAlter"/> command.
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
        Private Sub CmdLinkAlter_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdLinkAlter.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
               CodeEditor.IsCaretOnXmlBlock AndAlso
                CodeEditor.IsTextSelectedSingleLine()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdCollapse"/> command.
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
        Private Sub CmdCollapse_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdCollapse.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                (CodeEditor.TextContains(Documentation.XmlCommentCharsVB) OrElse
                CodeEditor.TextContains(Documentation.XmlCommentCharsCS))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdExpand"/> command.
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
        Private Sub CmdExpand_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdExpand.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                (CodeEditor.TextContains(Documentation.XmlCommentCharsVB) OrElse
                  CodeEditor.TextContains(Documentation.XmlCommentCharsCS))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdDelete"/> command.
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
        Private Sub CmdDelete_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdDelete.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                (CodeEditor.TextContains(Documentation.XmlCommentCharsVB) OrElse
                 CodeEditor.TextContains(Documentation.XmlCommentCharsCS))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdSeparator"/> command.
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
        Private Sub CmdSeparator_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdSeparator.BeforeQueryStatus

            ' Nothing to do here.

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdParagraph"/> command.
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
        Private Sub CmdParagraph_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdParagraph.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                CodeEditor.IsCaretOnXmlBlock() 'AndAlso CodeEditor.IsTextSelected()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdRemarks"/> command.
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
        Private Sub CmdRemarks_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdRemarks.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled =
                CodeEditor.IsCaretOnXmlBlock AndAlso
                CodeEditor.IsTextSelected()

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Handles the <see cref="OleMenuCommand.BeforeQueryStatus"/> event of the <see cref="CmdCreateSnippet"/> command.
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
        Private Sub CmdSnippet_BeforeQueryStatus(sender As Object, e As EventArgs) _
        Handles CmdCreateSnippet.BeforeQueryStatus

            DirectCast(sender, OleMenuCommand).Enabled = CodeEditor.IsTextSelected()

        End Sub

#End Region

    End Class

End Namespace
