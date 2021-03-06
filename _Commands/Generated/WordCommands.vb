Namespace Commands

	''' <summary></summary>
	''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
	''' <generator-date>17/02/2014 16:02:45</generator-date>
	''' <generator-functions>1</generator-functions>
	''' <generator-source>Deimos\_Commands\Generated\WordCommands.tt</generator-source>
	''' <generator-template>Text-Templates\Classes\VB_Object.tt</generator-template>
	''' <generator-version>1</generator-version>
	<System.CodeDom.Compiler.GeneratedCode("Deimos\_Commands\Generated\WordCommands.tt", "1")> _
	<Leviathan.Commands.Command(ResourceContainingType:=GetType(WordCommands), ResourceName:="CommandDetails", Name:="word", Description:="@commandOfficeDescriptionWord@", Hidden:=False)> _
	Partial Public Class WordCommands
		Inherits System.Object
		Implements System.IDisposable

		#Region " Public Constructors "

			''' <summary>Parametered Constructor (1 Parameters)</summary>
			Public Sub New( _
				ByVal _Host As Leviathan.Commands.ICommandsExecution _
			)

				MyBase.New()

				Host = _Host

			End Sub

			''' <summary>Parametered Constructor (2 Parameters)</summary>
			Public Sub New( _
				ByVal _Host As Leviathan.Commands.ICommandsExecution, _
				ByVal _Document As Deimos.Word.DocumentWrapper _
			)

				MyBase.New()

				Host = _Host
				Document = _Document

			End Sub

		#End Region

		#Region " Class Plumbing/Interface Code "

			#Region " IDisposable Implementation "

				#Region " Private Variables "

					''' <summary></summary>
					''' <remarks></remarks>
					Private IDisposable_DisposedCalled As System.Boolean

				#End Region

				#Region " Public Methods "

					Public Sub IDisposable_Dispose() Implements IDisposable.Dispose

						If Not IDisposable_DisposedCalled Then

							IDisposable_DisposedCalled = True
						End If

					End Sub

				#End Region

			#End Region

		#End Region

		#Region " Public Constants "

			''' <summary>Public Shared Reference to the Name of the Property: Host</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_HOST As String = "Host"

			''' <summary>Public Shared Reference to the Name of the Property: Document</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_DOCUMENT As String = "Document"

		#End Region

		#Region " Private Variables "

			''' <summary>Private Data Storage Variable for Property: Host</summary>
			''' <remarks></remarks>
			Private m_Host As Leviathan.Commands.ICommandsExecution

			''' <summary>Private Data Storage Variable for Property: Document</summary>
			''' <remarks></remarks>
			Private m_Document As Deimos.Word.DocumentWrapper

		#End Region

		#Region " Public Properties "

			''' <summary>Provides Access to the Property: Host</summary>
			''' <remarks></remarks>
			Public Property Host() As Leviathan.Commands.ICommandsExecution
				Get
					Return m_Host
				End Get
				Set(value As Leviathan.Commands.ICommandsExecution)
					m_Host = value
				End Set
			End Property

			''' <summary>Provides Access to the Property: Document</summary>
			''' <remarks></remarks>
			<Leviathan.Configuration.Configurable("name", ResourceContainingType:=GetType(WordCommands), ResourceName:="CommandDetails", Description:="@commandWordDescriptionName@", ArgsDescription:="@commandWordParameterDescriptionName@", Prefix:="/")> _
			Public Property Document() As Deimos.Word.DocumentWrapper
				Get
					Return m_Document
				End Get
				Set(value As Deimos.Word.DocumentWrapper)
					m_Document = value
				End Set
			End Property

		#End Region

	End Class

End Namespace