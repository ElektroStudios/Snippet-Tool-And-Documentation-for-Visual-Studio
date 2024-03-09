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
Imports EnvDTE
Imports EnvDTE80
Imports Microsoft.VisualStudio.Editor
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.TextManager.Interop

Imports ElektroStudios.SnippetToolPackage.MyPackage.Package
Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Formatting
Imports System.Text



#End Region
#Region " Code-Editor Tools "

Namespace MyPackage.Tools

    Public NotInheritable Class CodeEditor

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether there is a selected text in the code window editor.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <c>True</c> if there is selected text, <c>False</c> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function IsTextSelected() As Boolean

            Return Not GetCurrentViewHost().TextView.Selection.IsEmpty

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether there is a selected text within a single-line on the code window editor.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <c>True</c> if there is selected text within a single-line, <c>False</c> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function IsTextSelectedSingleLine() As Boolean

            Dim viewhost As IWpfTextViewHost = GetCurrentViewHost()
            Dim selection As ITextSelection = viewhost.TextView.Selection
            Dim isTextSelected As Boolean = Not selection.IsEmpty

            If isTextSelected Then
                ' Get the start and end points of the selection.
                Dim start As VirtualSnapshotPoint = selection.Start
                Dim [end] As VirtualSnapshotPoint = selection.End

                ' Get the lines that contain the start and end points.
                Dim startLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                Dim endLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                Dim isSameLine As Boolean = startLine.Equals(endLine)
                Return isSameLine
            End If

            Return False

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether there is a selected text within multiple lines on the code window editor.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <c>True</c> if there is selected text within a multiple lines, <c>False</c> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function IsTextSelectedMultiLine() As Boolean

            Dim viewhost As IWpfTextViewHost = GetCurrentViewHost()
            Dim selection As ITextSelection = viewhost.TextView.Selection
            Dim isTextSelected As Boolean = Not selection.IsEmpty

            If isTextSelected Then
                ' Get the start and end points of the selection.
                Dim start As VirtualSnapshotPoint = selection.Start
                Dim [end] As VirtualSnapshotPoint = selection.End

                ' Get the lines that contain the start and end points.
                Dim startLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition(start.Position)
                Dim endLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

                Dim isSameLine As Boolean = startLine.Equals(endLine)
                Return Not isSameLine
            End If

            Return False

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the caret is on a Xml documentation block on the code window editor.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <c>True</c> if there is on a Xml documentation block, <c>False</c> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function IsCaretOnXmlBlock() As Boolean

            Dim viewhost As IWpfTextViewHost = GetCurrentViewHost()

            ' Get the start and end points of the selection.
            Dim start As VirtualSnapshotPoint = viewhost.TextView.Selection.Start
            Dim [end] As VirtualSnapshotPoint = viewhost.TextView.Selection.End

            ' Get the lines that contain the start and end points.
            Dim startLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition(start.Position)
            Dim endLine As ITextViewLine = viewhost.TextView.GetTextViewLineContainingBufferPosition([end].Position.Subtract(1))

            ' Get the start and end points of the lines.
            Dim startLinePoint As SnapshotPoint = startLine.Start
            Dim endLinePoint As SnapshotPoint = endLine.End

            ' Create a SnapshotSpan for all text to be replaced.
            Dim span As New SnapshotSpan(startLinePoint, endLinePoint)
            Dim text As String = span.GetText?.TrimStart()
            Dim result As Boolean = Not String.IsNullOrWhiteSpace(text) AndAlso (text.StartsWith("'''") OrElse text.StartsWith("///"))
            Return result

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the source-code of the code window editor contains the specified text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="findText">
        ''' The text to find. 
        ''' </param>
        ''' 
        ''' <param name="ignoreCase">
        ''' If <c>True</c>, performs a ignore-case text-search. 
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <c>True</c> if contains the specified text, <c>False</c> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function TextContains(findText As String, Optional ignoreCase As Boolean = False) As Boolean

            Return If(ignoreCase, GetCurrentViewHost().TextView.TextSnapshot.GetText.ToLower.Contains(findText.ToLower),
                                  GetCurrentViewHost().TextView.TextSnapshot.GetText.Contains(findText))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the current view host.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see cref="IWpfTextViewHost"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function GetCurrentViewHost() As IWpfTextViewHost

            Dim txtMgr As IVsTextManager = DirectCast(Main.GetGlobalService(GetType(SVsTextManager)), IVsTextManager)
            Dim vTextView As IVsTextView = Nothing
            Dim mustHaveFocus As Integer = 1

            txtMgr.GetActiveView(mustHaveFocus, Nothing, vTextView)

            Dim userData As IVsUserData = TryCast(vTextView, IVsUserData)

            If userData Is Nothing Then
                Return Nothing

            Else
                Dim viewHost As IWpfTextViewHost
                Dim holder As Object = Nothing
                Dim guidViewHost As Guid = DefGuidList.guidIWpfTextViewHost
                userData.GetData(guidViewHost, holder)
                viewHost = DirectCast(holder, IWpfTextViewHost)
                Return viewHost

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the document language.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="dte">
        ''' The <see cref="EnvDTE80.DTE2"/> instance. 
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The language name like "BASIC" or "CSHARP". 
        ''' If there is any document open or else an open document without specific language, then it returns an empty string.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Function GetDocumentLanguage(dte As EnvDTE80.DTE2) As String

            If dte.ActiveWindow.Document IsNot Nothing Then

                Dim activeDoc As Document = dte.ActiveDocument
                Dim langString As String = activeDoc.Language
                Return If(String.IsNullOrEmpty(langString), "", langString)

            Else
                Return ""

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Deletes the Xml comments.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ce">
        ''' The member.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Sub DeleteComment(ce As CodeElement2)

            Dim memberStart As EditPoint
            Dim commentStart As EditPoint
            Dim commentEnd As EditPoint
            Dim comChars As String

            Select Case Main.Dte.ActiveDocument.ProjectItem.FileCodeModel.Language
                Case "{B5E9BD33-6D3E-4B5D-925E-8A43B79820B4}"
                    'VB
                    comChars = "'''"
                Case Else
                    'C#
                    comChars = "///"
            End Select

            memberStart = ce.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint
            commentStart = GetCommentStart(memberStart.CreateEditPoint, comChars)

            If (commentStart IsNot Nothing) Then
                commentEnd = GetCommentEnd(commentStart.CreateEditPoint, comChars)
                ' delete comment.
                commentStart.ReplaceText(commentEnd, "", 0)
            End If

            'try submembers
            If ce.IsCodeType Then
                Dim ce2 As CodeElement2
                For Each ce2 In CType(ce, CodeType).Members
                    DeleteComment(ce2)
                Next

            ElseIf ce.Kind = vsCMElement.vsCMElementNamespace Then
                Dim ce2 As CodeElement2
                For Each ce2 In CType(ce, CodeNamespace).Members
                    DeleteComment(ce2)
                Next

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Collapses the member and its sub members if any.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ce">
        ''' The member.
        ''' </param>
        ''' 
        ''' <param name="toggle">
        ''' If <see langword="True"/>, the comment outline is toggled, otherwise it is collapsed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Sub CollapseSubmembers(ce As CodeElement2, toggle As Boolean)

            Dim memberStart As EditPoint
            Dim commentStart As EditPoint
            Dim commentEnd As EditPoint
            Dim comChars As String

            Select Case Main.Dte.ActiveDocument.ProjectItem.FileCodeModel.Language
                Case "{B5E9BD33-6D3E-4B5D-925E-8A43B79820B4}"
                    'VB
                    comChars = "'''"
                Case Else
                    'C#
                    comChars = "///"
            End Select

            memberStart = ce.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes).CreateEditPoint
            commentStart = GetCommentStart(memberStart.CreateEditPoint, comChars)

            If (commentStart IsNot Nothing) Then
                commentEnd = GetCommentEnd(commentStart.CreateEditPoint, comChars)
                If toggle Then
                    'toggle
                    CType(Main.Dte.ActiveDocument.Selection, TextSelection).MoveToPoint(commentStart)
                    Main.Dte.ExecuteCommand("Edit.ToggleOutliningExpansion")
                Else
                    'collapse
                    commentStart.OutlineSection(commentEnd)
                End If
            End If

            'try submembers
            If ce.IsCodeType Then
                Dim ce2 As CodeElement2
                For Each ce2 In CType(ce, CodeType).Members
                    CollapseSubmembers(ce2, toggle)
                Next
            ElseIf ce.Kind = vsCMElement.vsCMElementNamespace Then
                Dim ce2 As CodeElement2
                For Each ce2 In CType(ce, CodeNamespace).Members
                    CollapseSubmembers(ce2, toggle)
                Next
            End If
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets starting point of the comment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ep">
        ''' Commented member start point.
        ''' </param>
        ''' 
        ''' <param name="commentChars">
        ''' The comment character. It is "'''" for VB or "///" for C#.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Function GetCommentStart(ep As EditPoint, commentChars As String) As EditPoint

            Try
                Dim line As String = ""
                Dim lastCommentLine As Integer

                ep.StartOfLine()
                ep.CharLeft()

                While Not ep.AtStartOfDocument
                    line = ep.GetLines(ep.Line, ep.Line + 1).Trim
                    If line.Length = 0 OrElse line.StartsWith(commentChars) Then
                        If line.Length > 0 Then
                            lastCommentLine = ep.Line
                        End If
                        ep.StartOfLine()
                        ep.CharLeft()
                    Else
                        Exit While
                    End If
                End While

                ep.MoveToLineAndOffset(lastCommentLine, 1)
                While ep.GetText(commentChars.Length) <> commentChars
                    ep.CharRight()
                End While

                Return ep.CreateEditPoint

            Catch ex As Exception
                Return Nothing

            End Try

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets ending point of the comment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="ep">
        ''' Comment start point.
        ''' </param>
        ''' 
        ''' <param name="commentChars">
        ''' The comment character. It is "'''" for VB or "///" for C#.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Private Shared Function GetCommentEnd(ep As EditPoint, commentChars As String) As EditPoint

            Try
                Dim line As String
                Dim lastCommentPoint As EditPoint
                lastCommentPoint = ep.CreateEditPoint
                ep.EndOfLine()
                ep.CharRight()

                While Not ep.AtEndOfDocument
                    line = ep.GetLines(ep.Line, ep.Line + 1).Trim
                    If line.StartsWith(commentChars) Then
                        lastCommentPoint = ep.CreateEditPoint
                        ep.EndOfLine()
                        ep.CharRight()
                    Else
                        Exit While
                    End If
                End While

                lastCommentPoint.EndOfLine()
                Return lastCommentPoint

            Catch ex As Exception
                Return Nothing

            End Try

        End Function

#End Region

    End Class

End Namespace

#End Region
