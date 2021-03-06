Namespace Visio

	''' <summary></summary>
	''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
	''' <generator-date>17/02/2014 16:06:23</generator-date>
	''' <generator-functions>1</generator-functions>
	''' <generator-source>Deimos\_Visio\Generated\DrawingWrapper.tt</generator-source>
	''' <generator-template>Text-Templates\Classes\VB_Object.tt</generator-template>
	''' <generator-version>1</generator-version>
	<System.CodeDom.Compiler.GeneratedCode("Deimos\_Visio\Generated\DrawingWrapper.tt", "1")> _
	Partial Public Class DrawingWrapper
		Inherits Deimos.DocumentBase

		#Region " Public Constructors "

			''' <summary>Default Constructor</summary>
			Public Sub New()

				MyBase.New()

				If Not Drawing Is Nothing Then Populate(Drawing, OfficeDocumentState.Existing, False, False)

			End Sub

			''' <summary>Parametered Constructor (1 Parameters)</summary>
			Public Sub New( _
				ByVal _Drawing As V.Document _
			)

				MyBase.New()

				m_Drawing = _Drawing

				If Not Drawing Is Nothing Then Populate(Drawing, OfficeDocumentState.Existing, False, False)

			End Sub

			''' <summary>Parametered Constructor (2 Parameters)</summary>
			Public Sub New( _
				ByVal _Drawing As V.Document, _
				ByVal _PageCount As System.Int32 _
			)

				MyBase.New()

				m_Drawing = _Drawing
				m_PageCount = _PageCount

				If Not Drawing Is Nothing Then Populate(Drawing, OfficeDocumentState.Existing, False, False)

			End Sub

		#End Region

		#Region " Public Constants "

			''' <summary>Public Shared Reference to the Name of the Property: Drawing</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_DRAWING As String = "Drawing"

			''' <summary>Public Shared Reference to the Name of the Property: PageCount</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_PAGECOUNT As String = "PageCount"

		#End Region

		#Region " Private Variables "

			''' <summary>Private Data Storage Variable for Property: Drawing</summary>
			''' <remarks></remarks>
			Private m_Drawing As V.Document

			''' <summary>Private Data Storage Variable for Property: PageCount</summary>
			''' <remarks></remarks>
			Private m_PageCount As System.Int32

		#End Region

		#Region " Public Properties "

			''' <summary>Provides Access to the Property: Drawing</summary>
			''' <remarks></remarks>
			Public ReadOnly Property Drawing() As V.Document
				Get
					Return m_Drawing
				End Get
			End Property

			''' <summary>Provides Access to the Property: PageCount</summary>
			''' <remarks></remarks>
			Public ReadOnly Property PageCount() As System.Int32
				Get
					Return m_PageCount
				End Get
			End Property

		#End Region

	End Class

End Namespace