''' <summary></summary>
''' <autogenerated>Generated from a T4 template. Modifications will be lost, if applicable use a partial class instead.</autogenerated>
''' <generator-date>17/02/2014 16:02:02</generator-date>
''' <generator-functions>1</generator-functions>
''' <generator-source>Deimos\General\Enums\OfficeDocumentState.tt</generator-source>
''' <generator-version>1</generator-version>
<System.CodeDom.Compiler.GeneratedCode("Deimos\General\Enums\OfficeDocumentState.tt", "1")> _
Public Enum OfficeDocumentState As System.Int32

	''' <summary>Indicates no State, e.g. no Document</summary>
	None = 0

	''' <summary>Indicates Document was already Open</summary>
	Existing = 1

	''' <summary>Indicates Document was Opened from File</summary>
	Opened = 2

	''' <summary>Indicates Document was Created</summary>
	Created = 3

End Enum