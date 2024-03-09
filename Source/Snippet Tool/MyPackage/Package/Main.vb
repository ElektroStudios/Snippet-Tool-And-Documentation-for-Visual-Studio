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

Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports EnvDTE80
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

Imports ElektroStudios.SnippetToolPackage.MyPackage.PackageSettings
Imports ElektroStudios.SnippetToolPackage.MyPackage.Tools
Imports ElektroStudios.SnippetToolPackage.MyPackage.UserInterface
Imports System.Threading
Imports System.Threading.Tasks


#End Region

Namespace MyPackage.Package

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' This is the class that implements the package exposed by this assembly.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' The <see cref="PackageRegistrationAttribute"/> attribute tells the 
    ''' PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
    ''' <para></para>
    ''' 
    ''' The <see cref="InstalledProductRegistrationAttribute"/> attribute is used to 
    ''' register the information needed to show this package in the Help/About dialog of Visual Studio.
    ''' <para></para>
    ''' 
    ''' The <see cref="ProvideMenuResourceAttribute"/> attribute is needed to 
    ''' let the shell know that this package exposes some menus.
    ''' <para></para>
    ''' 
    ''' The <see cref="ProvideAutoLoadAttribute"/> attribute is needed to auto-load the package when the specified condition occurs.
    ''' <para></para>
    ''' 
    ''' The <see cref="ProvideOptionPageAttribute"/> attribute is needed to provide the options page under "Tools -> Options" menu.
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    <InstalledProductRegistration("#110", "#112", "1.0", IconResourceID:=400)>
    <ProvideMenuResource("Menus.ctmenu", 1)>
    <ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)>
    <Guid(Guids.PackageString)>
    <ProvideOptionPage(GetType(SnippetsPageGrid), "Snippet Tool", "Snippet Defaults", 0, 0, True)>
    <ProvideOptionPage(GetType(XmlDocPageGrid), "Snippet Tool", "Xml Defaults", 0, 0, True)>
    <PackageRegistration(UseManagedResourcesOnly:=True, AllowsBackgroundLoading:=True)>
    Public NotInheritable Class Main : Inherits AsyncPackage

#Region " Private Fields "

#Region " DTE "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The <see cref="EnvDTE80.DTE2"/> instance. 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Shared Dte As DTE2

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The <see cref="DteInitializer"/> instance that initializes <see cref="Main.Dte"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend DteInitializer As DteInitializer

#End Region

#Region " Commands "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a Cref tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdCRef As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a Paramref tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdParamRef As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a langword tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdLangRef As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a single-line code tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdSinglelineCode As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a multi-line code tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdMultilineCode As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a code example tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdCodeExample As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a link tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdLink As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a link-alter tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdLinkAlter As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a Separator line.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdSeparator As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a Paragraph tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdParagraph As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes a remarks tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdRemarks As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that collapses Xml comments.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdCollapse As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that expands Xml comments.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdExpand As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that deletes Xml comments.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdDelete As OleMenuCommand

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The command that makes an Snippet.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private WithEvents CmdCreateSnippet As OleMenuCommand

#End Region

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="SnippetToolPackage"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New()

            ' Inside this method you can place any initialization code that does not require 
            ' any Visual Studio service because at this point the package object is created but 
            ' not sited yet inside Visual Studio environment. 

#If DEBUG Then
            ' Debug.WriteLine(String.Format("Entering constructor for: {0}", Me.GetType.Name))
#End If

        End Sub

#End Region

#Region " Overriden Methods "

        Protected Overrides Async Function InitializeAsync(cancellationToken As CancellationToken, progress As IProgress(Of ServiceProgressData)) As Task

#If DEBUG Then
            ' Debug.WriteLine(String.Format("Entering Initialize() of: {0}", Me.GetType().Name))
#End If
            ' When initialized asynchronously, the current thread may be a background thread at this point.
            ' Do any initialization that requires the UI thread after switching to the UI thread.
            Await Me.JoinableTaskFactory.SwitchToMainThreadAsync()

            Await MyBase.InitializeAsync(cancellationToken, progress)

            Me.InitializeDte()

            If (Main.Dte IsNot Nothing) Then
                Me.InitializeMenuHandlers()
            End If

        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sample code to get an instance of the <see cref="EnvDTE.DTE"/> root of the automation model.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://www.mztools.com/articles/2013/MZ2013029.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub InitializeDte()

            Dim shellService As IVsShell

            Main.Dte = TryCast(Me.GetService(GetType(SDTE)), DTE2)

            If Main.Dte Is Nothing Then
                ' The IDE is not yet fully initialized.
                shellService = TryCast(Me.GetService(GetType(SVsShell)), IVsShell)
                Me.DteInitializer = New DteInitializer(shellService, AddressOf Me.InitializeDte)

            Else
                Me.DteInitializer = Nothing

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initialized the menu handlers.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Sub InitializeMenuHandlers()

            ' Add our command handlers for menu (commands must exist in the .vsct file)
            Dim mcs As OleMenuCommandService = TryCast(Me.GetService(GetType(IMenuCommandService)), OleMenuCommandService)

            If mcs Is Nothing Then
                Thread.Sleep(15000)
                mcs = TryCast(Me.GetService(GetType(IMenuCommandService)), OleMenuCommandService)
            End If

            If mcs IsNot Nothing Then

                Dim cmdIdCRef As New CommandID(Guids.CmdSetRef, CommandIds.CRef)
                Me.CmdCRef = New OleMenuCommand(AddressOf Me.CmdCRef_Callback, cmdIdCRef)

                Dim cmdIdParamRef As New CommandID(Guids.CmdSetRef, CommandIds.ParamRef)
                Me.CmdParamRef = New OleMenuCommand(AddressOf Me.CmdParamRef_Callback, cmdIdParamRef)

                Dim cmdIdLangRef As New CommandID(Guids.CmdSetRef, CommandIds.LangRef)
                Me.CmdLangRef = New OleMenuCommand(AddressOf Me.CmdLangRef_Callback, cmdIdLangRef)

                Dim cmdIdSinglelineCode As New CommandID(Guids.CmdSetCode, CommandIds.SinglelineCode)
                Me.CmdSinglelineCode = New OleMenuCommand(AddressOf Me.CmdSinglelineCode_Callback, cmdIdSinglelineCode)

                Dim cmdIdMultilineCode As New CommandID(Guids.CmdSetCode, CommandIds.MultilineCode)
                Me.CmdMultilineCode = New OleMenuCommand(AddressOf Me.CmdMultilineCode_Callback, cmdIdMultilineCode)

                Dim cmdIdCodeExample As New CommandID(Guids.CmdSetCode, CommandIds.CodeExample)
                Me.CmdCodeExample = New OleMenuCommand(AddressOf Me.CmdCodeExample_Callback, cmdIdCodeExample)

                Dim cmdIdLink As New CommandID(Guids.CmdSetLink, CommandIds.Link)
                Me.CmdLink = New OleMenuCommand(AddressOf Me.CmdLink_Callback, cmdIdLink)

                Dim cmdIdLinkAlter As New CommandID(Guids.CmdSetLink, CommandIds.LinkAlter)
                Me.CmdLinkAlter = New OleMenuCommand(AddressOf Me.CmdLinkAlter_Callback, cmdIdLinkAlter)

                Dim cmdIdSeparator As New CommandID(Guids.CmdSetMisc, CommandIds.Separator)
                Me.CmdSeparator = New OleMenuCommand(AddressOf Me.CmdSeparator_Callback, cmdIdSeparator)

                Dim cmdIdParagraph As New CommandID(Guids.CmdSetMisc, CommandIds.Paragraph)
                Me.CmdParagraph = New OleMenuCommand(AddressOf Me.CmdParagraph_Callback, cmdIdParagraph)

                Dim cmdIdRemarks As New CommandID(Guids.CmdSetMisc, CommandIds.Remarks)
                Me.CmdRemarks = New OleMenuCommand(AddressOf Me.CmdRemarks_Callback, cmdIdRemarks)

                Dim cmdIdCollapse As New CommandID(Guids.CmdSetEditor, CommandIds.Collapse)
                Me.CmdCollapse = New OleMenuCommand(AddressOf Me.CmdCollapse_Callback, cmdIdCollapse)

                Dim cmdIdExpand As New CommandID(Guids.CmdSetEditor, CommandIds.Expand)
                Me.CmdExpand = New OleMenuCommand(AddressOf Me.CmdExpand_Callback, cmdIdExpand)

                Dim cmdIdDelete As New CommandID(Guids.CmdSetEditor, CommandIds.Delete)
                Me.CmdDelete = New OleMenuCommand(AddressOf Me.CmdDelete_Callback, cmdIdDelete)

                Dim cmdIdCreateSnippet As New CommandID(Guids.CmdSetSnippet, CommandIds.CreateSnippet)
                Me.CmdCreateSnippet = New OleMenuCommand(AddressOf Me.CmdCreateSnippet_Callback, cmdIdCreateSnippet)

                With mcs
                    .AddCommand(Me.CmdCRef)
                    .AddCommand(Me.CmdParamRef)
                    .AddCommand(Me.CmdLangRef)
                    .AddCommand(Me.CmdSinglelineCode)
                    .AddCommand(Me.CmdMultilineCode)
                    .AddCommand(Me.CmdCodeExample)
                    .AddCommand(Me.CmdLink)
                    .AddCommand(Me.CmdLinkAlter)
                    .AddCommand(Me.CmdSeparator)
                    .AddCommand(Me.CmdParagraph)
                    .AddCommand(Me.CmdRemarks)
                    .AddCommand(Me.CmdCollapse)
                    .AddCommand(Me.CmdExpand)
                    .AddCommand(Me.CmdDelete)
                    .AddCommand(Me.CmdCreateSnippet)
                End With

            End If

        End Sub

#End Region

#Region " IDisposable Implementation "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Releases the resources used by the <see cref="T:Microsoft.VisualStudio.Shell.Package"/> object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="disposing">
        ''' <c>True</c> if the object is being disposed, <c>False</c> if it is being finalized.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Protected Overrides Sub Dispose(disposing As Boolean)

            Main.Dte = Nothing
            MyBase.Dispose(disposing)

        End Sub

#End Region

    End Class

End Namespace
