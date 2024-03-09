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

#Region " XmlDoc PageGrid "

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
    Public NotInheritable Class XmlDocPageGrid : Inherits DialogPage

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the character used to fill the separator line.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The character used to fill the separator line.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Category("Separator")>
        <DisplayName("Character")>
        <Description("The character used to fill the separator line.")>
        Public Property Character() As Char = "-"c

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the separator line length.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The separator line length.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        <Category("Separator")>
        <DisplayName("Length")>
        <Description("The separator line length.")>
        Public Property Length() As Integer = 100

#End Region

    End Class

End Namespace

#End Region
