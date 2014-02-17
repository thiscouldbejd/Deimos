Imports System.Runtime.InteropServices
Imports System.Text

Namespace Media

	<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("203cffe3-2e18-4fdf-b59d-6e71530534cf")> _
	Public Interface IWMMetadataEditor2

		Sub Open( _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszFilename As String _
		)

		Sub Close()

		Sub Flush()

		Sub OpenEx(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszFilename As String, _
			<[In]()> ByVal dwDesiredAccess As UInt32, _
			<[In]()> ByVal dwShareMode As UInt32 _
		)

	End Interface

End Namespace