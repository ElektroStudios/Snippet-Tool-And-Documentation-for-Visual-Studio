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
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Imports Microsoft.VisualStudio.Shell

#End Region

#Region " Snippets PageGrid "

Namespace MyPackage.UserInterface

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' The dialog page grid of the Snippet options.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="https://msdn.microsoft.com/en-us/library/bb166195.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <DesignerCategory("Code")>
    <ClassInterface(ClassInterfaceType.AutoDual)>
    <CLSCompliant(False), ComVisible(True)>
    Public NotInheritable Class SnippetsPageGrid : Inherits DialogPage

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the author name for the snippets.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The author name for the snippets.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Category("Fields")>
        <DisplayName("Author")>
        <Description("The author name for the snippet generation.")>
        Public Property Author() As String = "Author"

#End Region

    End Class

End Namespace

#End Region
