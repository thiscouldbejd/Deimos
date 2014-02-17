Imports System.Runtime.InteropServices
Imports System.Text

Namespace Media

	<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("15CC68E3-27CC-4ECD-B222-3F5D02D80BD5")> _
	Public Interface IWMHeaderInfo3

		Function GetAttributeCount( _
			<[In]()> ByVal wStreamNum As UShort, _
			<Out()> ByRef pcAttributes As UShort _
		) As UInteger

		Function GetAttributeByIndex( _
			<[In]()> ByVal wIndex As UShort, _
			<Out(), [In]()> ByRef pwStreamNum As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pwszName As Byte(), _
			<Out(), [In]()> ByRef pcchNameLen As UShort, _
			<Out()> ByRef pType As WMT_ATTR_DATATYPE, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<Out(), [In]()> ByRef pcbLength As UShort _
		) As UInteger

		Function GetAttributeByName( _
			<Out(), [In]()> ByRef pwStreamNum As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pszName As Byte(), _
			<Out()> ByRef pType As WMT_ATTR_DATATYPE, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<Out(), [In]()> ByRef pcbLength As UShort _
		) As UInteger

		Function SetAttribute( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pszName As Byte(), _
			<[In]()> ByVal Type As WMT_ATTR_DATATYPE, _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<[In]()> ByVal cbLength As UShort _
		) As UInteger

		Function GetMarkerCount( _
			<Out()> ByRef pcMarkers As UShort _
		) As UInteger

		Function GetMarker( _
			<[In]()> ByVal wIndex As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pwszMarkerName As Byte(), _
			<Out(), [In]()> ByRef pcchMarkerNameLen As UShort, _
			<Out()> ByRef pcnsMarkerTime As ULong _
		) As UInteger

		Function AddMarker( _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pwszMarkerName As Byte(), _
			<[In]()> ByVal cnsMarkerTime As ULong _
		) As UInteger

		Function RemoveMarker( _
			<[In]()> ByVal wIndex As UShort _
		) As UInteger

		Function GetScriptCount( _
			<Out()> ByRef pcScripts As UShort _
		) As UInteger

		Function GetScript( _
			<[In]()> ByVal wIndex As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszType As String, _
			<Out(), [In]()> ByRef pcchTypeLen As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszCommand As String, _
			<Out(), [In]()> ByRef pcchCommandLen As UShort, _
			<Out()> ByRef pcnsScriptTime As ULong _
		) As UInteger

		Function AddScript( _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszType As String, _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszCommand As String, _
			<[In]()> ByVal cnsScriptTime As ULong _
		) As UInteger

		Function RemoveScript( _
			<[In]()> ByVal wIndex As UShort _
		) As UInteger

		Function GetCodecInfoCount( _
			<Out()> ByRef pcCodecInfos As UInteger _
		) As UInteger

		Function GetCodecInfo( _
			<[In]()> ByVal wIndex As UInteger, _
			<Out(), [In]()> ByRef pcchName As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszName As String, _
			<Out(), [In]()> ByRef pcchDescription As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszDescription As String, _
			<Out()> ByRef pCodecType As WMT_CODEC_INFO_TYPE, _
			<Out(), [In]()> ByRef pcbCodecInfo As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pbCodecInfo As Byte() _
		) As UInteger

		Function GetAttributeCountEx( _
			<[In]()> ByVal wStreamNum As UShort, _
			<Out()> ByRef pcAttributes As UShort _
		) As UInteger

		Function GetAttributeIndices( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszName As String, _
			<[In]()> ByRef pwLangIndex As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pwIndices As UShort(), _
			<Out(), [In]()> ByRef pwCount As UShort _
		) As UInteger

		Function GetAttributeByIndexEx( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In]()> ByVal wIndex As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszName As String, _
			<Out(), [In]()> ByRef pwNameLen As UShort, _
			<Out()> ByRef pType As WMT_ATTR_DATATYPE, _
			<Out()> ByRef pwLangIndex As UShort, _
			<Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<Out(), [In]()> ByRef pdwDataLength As UInteger _
		) As UInteger

		Function ModifyAttribute( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In]()> ByVal wIndex As UShort, _
			<[In]()> ByVal Type As WMT_ATTR_DATATYPE, _
			<[In]()> ByVal wLangIndex As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<[In]()> ByVal dwLength As UInteger _
		) As UInteger

		Function AddAttribute( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String, _
			<Out()> ByRef pwIndex As UShort, _
			<[In]()> ByVal Type As WMT_ATTR_DATATYPE, _
			<[In]()> ByVal wLangIndex As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pValue As Byte(), _
			<[In]()> ByVal dwLength As UInteger _
		) As UInteger

		Function DeleteAttribute( _
			<[In]()> ByVal wStreamNum As UShort, _
			<[In]()> ByVal wIndex As UShort _
		) As UInteger

		Function AddCodecInfo( _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String, _
			<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pwszDescription As String, _
			<[In]()> ByVal codecType As WMT_CODEC_INFO_TYPE, _
			<[In]()> ByVal cbCodecInfo As UShort, _
			<[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pbCodecInfo As Byte() _
		) As UInteger

	End Interface

End Namespace