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

#End Region

#Region " Documentation "

Namespace MyPackage.Tools

    Partial Public NotInheritable Class Documentation

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The characters used to identify an XML comment in VisualBasic.Net language.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const XmlCommentCharsVB As String = "'''"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The characters used to identify an XML comment in C-Sharp language.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const XmlCommentCharsCS As String = "///"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The snippet template for VisualBasic.Net language.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly SnippetTemplateFormatVB As String =
<a><![CDATA[<?xml version="1.0" encoding="utf-8"?>
        <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
        <CodeSnippet Format="1.0.0">

        <Header>
          <Title>Title</Title>
          <Description>Description</Description>
          <Author>{0}</Author>
        </Header>

        <Snippet>

        <References>
          <Reference>
            <Assembly>System.dll</Assembly>
          </Reference>
        </References>

        <Imports>
          <Import>
            <Namespace>System</Namespace>
          </Import>
        </Imports>

        <Declarations>
          <Literal Editable="false">
            <ID>CDATAEnd</ID>
            <ToolTip>CDATA end tag to scape XML illegal characters, if needed.</ToolTip>
            <Default>&gt;</Default>
          </Literal>
        </Declarations>

        <Code Language="VB"><![CDATA[

        {1}

        ]]$cdataend$</a>
          </Snippet>
         </CodeSnippet>
        </CodeSnippets>]]></a>.Value

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The snippet template for C-Sharp language.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly SnippetTemplateFormatCS As String =
<a><![CDATA[<?xml version="1.0" encoding="utf-8"?>
        <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
        <CodeSnippet Format="1.0.0">

        <Header>
          <Title>Title</Title>
          <Description>Description</Description>
          <Author>{0}</Author>
        </Header>

        <Snippet>

        <References>
          <Reference>
            <Assembly>System.dll</Assembly>
          </Reference>
        </References>

        <Declarations>
          <Literal Editable="false">
            <ID>CDATAEnd</ID>
            <ToolTip>CDATA end tag to scape XML illegal characters, if needed.</ToolTip>
            <Default>&gt;</Default>
          </Literal>
        </Declarations>

        <Code Language="CSharp"><![CDATA[

        {1}

        ]]$cdataend$</a>
          </Snippet>
         </CodeSnippet>
        </CodeSnippets>]]></a>.Value

    End Class

End Namespace

#End Region
