Namespace Excel

	''' <summary></summary>
	''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
	''' <generator-date>17/02/2014 16:02:51</generator-date>
	''' <generator-functions>1</generator-functions>
	''' <generator-source>Deimos\_Excel\Generated\WorkbookWrapper.tt</generator-source>
	''' <generator-template>Text-Templates\Classes\VB_Object.tt</generator-template>
	''' <generator-version>1</generator-version>
	<System.CodeDom.Compiler.GeneratedCode("Deimos\_Excel\Generated\WorkbookWrapper.tt", "1")> _
	Partial Public Class WorkbookWrapper
		Inherits Deimos.DocumentBase

		#Region " Public Constructors "

			''' <summary>Default Constructor</summary>
			Public Sub New()

				MyBase.New()

				If Not Workbook Is Nothing Then Populate(Workbook, OfficeDocumentState.Existing, False, False)

			End Sub

			''' <summary>Parametered Constructor (1 Parameters)</summary>
			Public Sub New( _
				ByVal _Workbook As E.Workbook _
			)

				MyBase.New()

				m_Workbook = _Workbook

				If Not Workbook Is Nothing Then Populate(Workbook, OfficeDocumentState.Existing, False, False)

			End Sub

			''' <summary>Parametered Constructor (2 Parameters)</summary>
			Public Sub New( _
				ByVal _Workbook As E.Workbook, _
				ByVal _SheetCount As System.Int32 _
			)

				MyBase.New()

				m_Workbook = _Workbook
				m_SheetCount = _SheetCount

				If Not Workbook Is Nothing Then Populate(Workbook, OfficeDocumentState.Existing, False, False)

			End Sub

		#End Region

		#Region " Public Constants "

			''' <summary>Public Shared Reference to the Name of the Property: Workbook</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_WORKBOOK As String = "Workbook"

			''' <summary>Public Shared Reference to the Name of the Property: SheetCount</summary>
			''' <remarks></remarks>
			Public Const PROPERTY_SHEETCOUNT As String = "SheetCount"

		#End Region

		#Region " Private Variables "

			''' <summary>Private Data Storage Variable for Property: Workbook</summary>
			''' <remarks></remarks>
			Private m_Workbook As E.Workbook

			''' <summary>Private Data Storage Variable for Property: SheetCount</summary>
			''' <remarks></remarks>
			Private m_SheetCount As System.Int32

		#End Region

		#Region " Public Properties "

			''' <summary>Provides Access to the Property: Workbook</summary>
			''' <remarks></remarks>
			Public ReadOnly Property Workbook() As E.Workbook
				Get
					Return m_Workbook
				End Get
			End Property

			''' <summary>Provides Access to the Property: SheetCount</summary>
			''' <remarks></remarks>
			Public ReadOnly Property SheetCount() As System.Int32
				Get
					Return m_SheetCount
				End Get
			End Property

		#End Region

	End Class

End Namespace