''' <summary></summary>
''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
''' <generator-date>17/02/2014 16:02:09</generator-date>
''' <generator-functions>1</generator-functions>
''' <generator-source>Deimos\General\Generated\PageBase.tt</generator-source>
''' <generator-template>Text-Templates\Classes\VB_Object.tt</generator-template>
''' <generator-version>1</generator-version>
<System.CodeDom.Compiler.GeneratedCode("Deimos\General\Generated\PageBase.tt", "1")> _
Partial Public MustInherit Class PageBase
	Inherits System.Object

	#Region " Public Constructors "

		''' <summary>Default Constructor</summary>
		Public Sub New()

			MyBase.New()

			m_State = Deimos.OfficePageState.None
		End Sub

		''' <summary>Parametered Constructor (1 Parameters)</summary>
		Public Sub New( _
			ByVal _Parent As Deimos.DocumentBase _
		)

			MyBase.New()

			Parent = _Parent

			m_State = Deimos.OfficePageState.None
		End Sub

		''' <summary>Parametered Constructor (2 Parameters)</summary>
		Public Sub New( _
			ByVal _Parent As Deimos.DocumentBase, _
			ByVal _State As Deimos.OfficePageState _
		)

			MyBase.New()

			Parent = _Parent
			m_State = _State

		End Sub

	#End Region

	#Region " Public Constants "

		''' <summary>Public Shared Reference to the Name of the Property: Parent</summary>
		''' <remarks></remarks>
		Public Const PROPERTY_PARENT As String = "Parent"

		''' <summary>Public Shared Reference to the Name of the Property: State</summary>
		''' <remarks></remarks>
		Public Const PROPERTY_STATE As String = "State"

	#End Region

	#Region " Private Variables "

		''' <summary>Private Data Storage Variable for Property: Parent</summary>
		''' <remarks></remarks>
		Private m_Parent As Deimos.DocumentBase

		''' <summary>Private Data Storage Variable for Property: State</summary>
		''' <remarks></remarks>
		Private m_State As Deimos.OfficePageState

	#End Region

	#Region " Public Properties "

		''' <summary>Provides Access to the Property: Parent</summary>
		''' <remarks></remarks>
		Public Overridable Property Parent() As Deimos.DocumentBase
			Get
				Return m_Parent
			End Get
			Set(value As Deimos.DocumentBase)
				m_Parent = value
			End Set
		End Property

		''' <summary>Provides Access to the Property: State</summary>
		''' <remarks></remarks>
		Public ReadOnly Property State() As Deimos.OfficePageState
			Get
				Return m_State
			End Get
		End Property

	#End Region

End Class