' ***********************************************************************
' Author   : Elektro
' Modified : 25-June-2022
' ***********************************************************************

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports EnvDTE80
Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Formatting

Imports ElektroStudios.SnippetToolPackage.MyPackage.Enums
Imports ElektroStudios.SnippetToolPackage.MyPackage.Package

#End Region

Namespace MyPackage.Tools

    Public NotInheritable Class Documentation

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Documents the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="task">
        ''' The documentation task to perform.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Zero if success, non-zero if fails.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function ProcessSelectedText(dte As DTE2, viewHost As IWpfTextViewHost, task As DocumentationTask) As Integer

            Try

                Dim textView As IWpfTextView = viewHost.TextView
                Dim selection As ITextSelection = textView.Selection
                Dim language As String = CodeEditor.GetDocumentLanguage(dte)

                Dim codeLang As String = ""
                Dim xmlCommentChars As String = ""
                Select Case language.ToUpper()

                    Case "BASIC"
                        xmlCommentChars = XmlCommentCharsVB
                        codeLang = "VB"

                    Case "CSHARP"
                        xmlCommentChars = XmlCommentCharsCS
                        codeLang = "CSharp"

                    Case ""
                        ' Any document is open or else document without specific language.
                        Return -1

                    Case Else ' VC++
                        Return -1

                End Select

                ' Construct the replacement string.
                Select Case task

                    Case DocumentationTask.MakeCodeRef
                        DoTask_MakeCodeRef(selection)

                    Case DocumentationTask.MakeParamRef
                        DoTask_MakeParamRef(selection)

                    Case DocumentationTask.MakeLangRef
                        DoTask_MakeLangRef(selection)

                    Case DocumentationTask.MakeSinglelineCode
                        DoTask_MakeSinglelineCode(selection)

                    Case DocumentationTask.MakeMultilineCode
                        DoTask_MakeMultilineCode(selection)

                    Case DocumentationTask.MakeCodeExample
                        DoTask_MakeCodeExample(selection, xmlCommentChars, codeLang)

                    Case DocumentationTask.MakeLink
                        DoTask_MakeLink(selection)

                    Case DocumentationTask.MakeLinkAlter
                        DoTask_MakeLinkAlter(selection)

                    Case DocumentationTask.MakeSeparator
                        DoTask_MakeSeparator(selection, xmlCommentChars)

                    Case DocumentationTask.MakeParagraph
                        DoTask_MakeParagraph(selection, xmlCommentChars)

                    Case DocumentationTask.MakeRemarks
                        DoTask_MakeRemarks(selection, xmlCommentChars)

                    Case DocumentationTask.CollapseXmlComments
                        DoTask_CollapseXmlComments()

                    Case DocumentationTask.ExpandXmlComments
                        DoTask_ExpandXmlComments()

                    Case DocumentationTask.DeleteXmlComments
                        DoTask_DeleteXmlComments()

                    Case DocumentationTask.MakeSnippet
                        DoTask_MakeSnippet(selection, language)

                    Case Else
                        Return -1

                End Select

                Return 0

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Snippet Tool")
                Return -1

            End Try

        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeCodeRef"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeCodeRef(selection As ITextSelection)

            If Not CodeEditor.IsTextSelected() Then
                Main.Dte.ExecuteCommand("Edit.SelectCurrentWord")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see cref=""{0}""/>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeParamRef"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeParamRef(selection As ITextSelection)

            If Not CodeEditor.IsTextSelected() Then
                Main.Dte.ExecuteCommand("Edit.SelectCurrentWord")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<paramref name=""{0}""/>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeLangRef"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeLangRef(selection As ITextSelection)

            If Not CodeEditor.IsTextSelected() Then
                Main.Dte.ExecuteCommand("Edit.SelectCurrentWord")
            End If

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see langword=""{0}""/>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeSinglelineCode"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeSinglelineCode(selection As ITextSelection)

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<c>{0}</c>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeMultilineCode"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeMultilineCode(selection As ITextSelection)

            Try
                Main.Dte.UndoContext.Open("Make multiline code")

                ' Create a SnapshotSpan for all text to be replaced.
                Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

                ' Perform the replacement.
                span.Snapshot.TextBuffer.Replace(span, String.Format("<code>{0}</code>", span.GetText))
                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeCodeExample"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="xmlCommentChars">
        ''' The XML comment chars to use.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeCodeExample(selection As ITextSelection, xmlCommentChars As String, codeLang As String)

            Try
                Main.Dte.UndoContext.Open("Make code example")

                ' Get the start and end points of the selection.
                Dim start As VirtualSnapshotPoint = selection.Start
                Dim [end] As VirtualSnapshotPoint = selection.End

                ' Get the lines that contain the start and end points.
                Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                ' Get the start and end points of the lines.
                Dim startLinePoint As SnapshotPoint = startLine.Start
                Dim endLinePoint As SnapshotPoint = endLine.End

                ' Create a SnapshotSpan for all text to be replaced.
                Dim span As New SnapshotSpan(startLinePoint, endLinePoint)

                ' Compute margin.
                Dim lines As IEnumerable(Of String) =
                    span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None) '.SkipWhile(Function(line As String) String.IsNullOrWhiteSpace(line))

                Dim margin As Integer = lines.
                    Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                    Select(Function(line)
                               Dim count As Integer = 0
                               While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                               End While
                               Return Interlocked.Decrement(count)
                           End Function).Min

                Dim sb As New StringBuilder

                With sb
                    .AppendLine(String.Format("{0} <example> This is a code example.", xmlCommentChars))
                    .AppendLine(String.Format("{0} <code language=""{1}"">", xmlCommentChars, codeLang))

                    For Each line As String In lines

                        If String.IsNullOrWhiteSpace(line) Then
                            sb.AppendLine(New String(" "c, margin) & String.Format("{0} ", xmlCommentChars))

                        Else
                            sb.AppendLine(String.Format("{0}{1}{2}{3}", New String(" "c, margin), xmlCommentChars, If(margin = 0, " ", ""), line.Remove(0, margin)))

                        End If

                    Next line

                    .AppendLine(String.Format("{0} </code>", xmlCommentChars))
                    .AppendLine(String.Format("{0} </example>", xmlCommentChars))
                End With

                ' Perform the replacement.
                span.Snapshot.TextBuffer.Replace(span, sb.ToString)
                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeLink"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeLink(selection As ITextSelection)

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<see href=""{0}""/>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeLinkAlter"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeLinkAlter(selection As ITextSelection)

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

            ' Perform the replacement.
            span.Snapshot.TextBuffer.Replace(span, String.Format("<seealso href=""{0}""/>", span.GetText))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeSeparator"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeSeparator(selection As ITextSelection, xmlCommentChars As String)

            Try
                Main.Dte.UndoContext.Open("Make separator line")

                Dim c As Char =
                    Convert.ToChar(Main.Dte.Properties("Snippet Tool", "Xml Defaults").Item("Character").Value)

                Dim length As Integer =
                   CInt(Main.Dte.Properties("Snippet Tool", "Xml Defaults").Item("Length").Value)

                Select Case CodeEditor.IsTextSelected()

                    Case False
                        Dim caret As ITextCaret = CodeEditor.GetCurrentViewHost.TextView.Caret
                        Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                        Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                        Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                        Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                        snapshot.TextBuffer.Replace(span, line &
                                                          ControlChars.NewLine &
                                                          New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)))

                    Case Else

                        ' Get the start and end points of the selection.
                        Dim start As VirtualSnapshotPoint = selection.Start
                        Dim [end] As VirtualSnapshotPoint = selection.End

                        ' Get the lines that contain the start and end points.
                        Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                        Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                        ' Get the start and end points of the lines.
                        Dim startLinePoint As SnapshotPoint = startLine.Start
                        Dim endLinePoint As SnapshotPoint = endLine.End

                        ' Create a SnapshotSpan for all text to be replaced.
                        Dim span As New SnapshotSpan(startLinePoint, endLinePoint)

                        Dim lines As IEnumerable(Of String) =
                            span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

                        ' Compute margin.
                        Dim margin As Integer = lines.
                                       Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                                       Select(Function(line)
                                                  Dim count As Integer = 0
                                                  While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                                                  End While
                                                  Return Interlocked.Decrement(count)
                                              End Function).Min

                        ' Perform the replacement.
                        span.Snapshot.TextBuffer.Replace(span, New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)) &
                                                               ControlChars.NewLine &
                                                               span.GetText &
                                                               ControlChars.NewLine &
                                                               New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, New String(c, length)))

                End Select

                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeParagraph"/> task on the selected text.
        ''' <para></para>
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeParagraph(selection As ITextSelection, xmlCommentChars As String)

            Try
                Main.Dte.UndoContext.Open("Make paragraph")

                If CodeEditor.IsTextSelected() Then
                    ' Create a SnapshotSpan for all text to be replaced.
                    Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

                    ' Perform the replacement.
                    span.Snapshot.TextBuffer.Replace(span, String.Format("<para>{0}</para>", span.GetText))

                Else
                    Dim caret As ITextCaret = CodeEditor.GetCurrentViewHost.TextView.Caret
                    Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                    Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                    Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                    Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                    snapshot.TextBuffer.Replace(span, line & ControlChars.NewLine & New String(" "c, margin) &
                                                String.Format("{0} {1}", xmlCommentChars, "<para></para>"))
                End If

                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeRemarks"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="xmlCommentChars">
        ''' The XML comment chars to use.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeRemarks(selection As ITextSelection, xmlCommentChars As String)

            Try
                Main.Dte.UndoContext.Open("Make remarks")

                Select Case CodeEditor.IsTextSelected()

                    Case False
                        Dim caret As ITextCaret = CodeEditor.GetCurrentViewHost.TextView.Caret
                        Dim span As Span = caret.ContainingTextViewLine.Extent.Span
                        Dim snapshot As ITextSnapshot = caret.ContainingTextViewLine.Extent.Snapshot
                        Dim line As String = snapshot.GetLineFromPosition(caret.Position.BufferPosition.Position).GetText
                        Dim margin As Integer = line.TakeWhile(Function(chr As Char) Char.IsWhiteSpace(chr)).Count

                        snapshot.TextBuffer.Replace(span, line &
                                                          ControlChars.NewLine &
                                                          New String(" "c, margin) & String.Format("{0} {1}", xmlCommentChars, "<remarks></remarks>"))

                    Case Else
                        ' Create a SnapshotSpan for all text to be replaced.
                        Dim span As SnapshotSpan = selection.StreamSelectionSpan.SnapshotSpan

                        Dim lines As IEnumerable(Of String) =
                            span.GetText.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

                        ' Compute margin.
                        Dim margin As Integer = lines.
                                       Where(Function(line As String) Not String.IsNullOrWhiteSpace(line)).
                                       Select(Function(line)
                                                  Dim count As Integer = 0
                                                  While Char.IsWhiteSpace(line(Math.Max(Interlocked.Increment(count), count - 1)))
                                                  End While
                                                  Return Interlocked.Decrement(count)
                                              End Function).Min

                        Dim sb As New StringBuilder
                        With sb
                            .AppendLine(String.Format("{0} <remarks>", xmlCommentChars))

                            For Each line As String In lines

                                If String.IsNullOrWhiteSpace(line) Then
                                    sb.AppendLine(New String(" "c, margin) & String.Format("{0} ", xmlCommentChars))

                                Else
                                    sb.AppendLine(String.Format("{0}{1}{2}{3}", New String(" "c, margin), xmlCommentChars, If(margin = 0, " ", ""), line.Remove(0, margin)))

                                End If

                            Next line

                            .AppendLine(String.Format("{0} </remarks>", xmlCommentChars))
                        End With

                        ' Perform the replacement.
                        span.Snapshot.TextBuffer.Replace(span, sb.ToString)

                End Select

                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.CollapseXmlComments"/> task.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_CollapseXmlComments()

            Try
                Main.Dte.UndoContext.Open("Collapse XML comments")

                For Each ce As CodeElement2 In Main.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                    CodeEditor.CollapseSubmembers(ce, False)
                Next

                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.ExpandXmlComments"/> task.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_ExpandXmlComments()

            Try
                Main.Dte.UndoContext.Open("Expand XML comments")

                For Each ce As CodeElement2 In Main.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                    CodeEditor.CollapseSubmembers(ce, False)
                    CodeEditor.CollapseSubmembers(ce, True)
                Next

                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.DeleteXmlComments"/> task.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_DeleteXmlComments()

            Try
                Main.Dte.UndoContext.Open("Delete XML comments")
                For Each ce As CodeElement2 In Main.Dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements
                    CodeEditor.DeleteComment(ce)
                Next
                Main.Dte.UndoContext.Close()

            Catch ex As Exception
                Main.Dte.UndoContext.Close()

            End Try

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs the <see cref="DocumentationTask.MakeSnippet"/> task on the selected text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="selection">
        ''' The <see cref="ITextSelection"/> instance.
        ''' </param>
        ''' 
        ''' <param name="languageName">
        ''' The current document language name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Sub DoTask_MakeSnippet(selection As ITextSelection, languageName As String)

            ' Get the start and end points of the selection.
            Dim start As VirtualSnapshotPoint = selection.Start
            Dim [end] As VirtualSnapshotPoint = selection.End

            ' Get the lines that contain the start and end points.
            Dim startLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition(start.Position)
            Dim endLine As ITextViewLine = selection.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

            ' Get the start and end points of the lines.
            Dim startLinePoint As SnapshotPoint = startLine.Start
            Dim endLinePoint As SnapshotPoint = endLine.End

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As New SnapshotSpan(startLinePoint, endLinePoint)

            Dim author As String =
                Main.Dte.Properties("Snippet Tool", "Snippet Defaults").Item("Author").Value.ToString

            Dim sb As New StringBuilder
            Select Case languageName.ToUpper

                Case "BASIC"
                    sb.AppendFormat(SnippetTemplateFormatVB, author, span.GetText).Replace("$cdataend$", ">")

                Case "CSHARP"
                    sb.AppendFormat(SnippetTemplateFormatCS, author, span.GetText).Replace("$cdataend$", ">")

                Case Else

            End Select

            Dim tempFileName As String = String.Format("{0}.snippet", Path.GetTempFileName)
            File.WriteAllText(tempFileName, sb.ToString, Encoding.Default)
            Diagnostics.Process.Start(tempFileName)

        End Sub

#End Region

    End Class

End Namespace
