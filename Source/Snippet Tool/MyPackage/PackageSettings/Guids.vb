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

#End Region

#Region " Guids "

Namespace MyPackage.PackageSettings

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Exposes the package info, such as GUIDs and Command Identifiers.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Friend Module Guids

#Region " Package "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of this package, represented in String datatype.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Const PackageString As String = "8da2cd63-a9eb-488b-bd47-75f83b617766"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of this package.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly Package As New Guid(Guids.PackageString)

#End Region

#Region " Command-sets "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Ref" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetRef As New Guid("cac498f9-5603-44f9-b038-bd8223793060")

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Code" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetCode As New Guid("cac498f9-5603-44f9-b038-bd8223793061")

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Link" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetLink As New Guid("cac498f9-5603-44f9-b038-bd8223793062")

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Misc" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetMisc As New Guid("cac498f9-5603-44f9-b038-bd8223793063")

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Misc" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetEditor As New Guid("cac498f9-5603-44f9-b038-bd8223793064")

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The unique identifier of "Snippet" command-set.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend ReadOnly CmdSetSnippet As New Guid("cac498f9-5603-44f9-b038-bd8223793065")

#End Region

    End Module

End Namespace

#End Region
