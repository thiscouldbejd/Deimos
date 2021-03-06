Namespace Word

	''' <summary></summary>
	''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
	''' <generator-date>17/02/2014 16:06:37</generator-date>
	''' <generator-functions>1</generator-functions>
	''' <generator-source>Deimos\_Word\Generated\DocumentWrapper.tt</generator-source>
	''' <generator-template>Text-Templates\Classes\VB_Object.tt</generator-template>
	''' <generator-version>1</generator-version>
	<System.CodeDom.Compiler.GeneratedCode("Deimos\_Word\Generated\DocumentWrapper.tt", "1")> _
	Partial Public Class DocumentWrapper
		Inherits Deimos.DocumentBase

		#Region " Public Constructors "

			''' <summary>Default Constructor</summary>
			Public Sub New()

				MyBase.New()

				If Not Document Is Nothing Then Populate(Document, OfficeDocumentState.Existing, False, False)

			End Sub

			''' <summary>Parametered Constructor (1 Parameters)</summary>
			Public Sub New( _
				ByVal _Document As W.Document _
			)

				MyBase.New()

				m_Document = _Document

				If Not Document Is Nothing Then Populate(Document, OfficeDocumentState.Existing, False, False)

			End Sub

		#End Region

		#Region " Public Constants "

			''' <summary>Public Shared Reference to the Name of the Property: Document</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_DOCUMENT As String = "Document"

		#End Region

		#Region " Private Variables "

			''' <summary>Private Data Storage Variable for Property: Document</summary>
			''' <remarks></remarks>
			Private m_Document As W.Document

		#End Region

		#Region " Public Properties "

			''' <summary>Provides Access to the Property: Document</summary>
			''' <remarks></remarks>
			Public ReadOnly Property Document() As W.Document
				Get
					Return m_Document
				End Get
			End Property

		#End Region

	End Class

End Namespace