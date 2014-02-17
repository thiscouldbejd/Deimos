Imports Deimos.Media
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Commands

	Partial Public Class MediaCommands

		#Region " Private Methods "

			Private Function SetAttribute( _
				ByVal file As FileInfo, _
				ByVal attribute As String, _
				ByVal value As String, _
				Optional ByVal stream As UInt16 = UInt16.MaxValue _
			) As Boolean

				Dim metadataEditor As IWMMetadataEditor2 = Nothing

				Try

					MediaCommands.CreateEditor(metadataEditor)

					metadataEditor.Open(file.FullName)

					Dim attribute_Type As WMT_ATTR_DATATYPE
					Dim attribute_Name As Byte() = GetBytes(attribute, WMT_ATTR_DATATYPE.WMT_TYPE_STRING)
					Dim attribute_Value As Byte() = Nothing
					Dim attribute_ValueLength As UInt16

					CType(metadataEditor, IWMHeaderInfo3).GetAttributeByName( _
					stream, attribute_Name, attribute_Type, _
					attribute_Value, attribute_ValueLength)

					attribute_Value = GetBytes(value, attribute_Type)
					attribute_ValueLength = attribute_Value.Length

					CType(metadataEditor, IWMHeaderInfo3).SetAttribute(stream, attribute_Name, _
					attribute_Type, attribute_Value, value.Length * 2)

				Catch cex As COMException

					If (cex.ErrorCode = -1072889827) Then Throw New FileNotFoundException( _
						"Failed to open the file into memory. The file may  be missing or you may not have permission to open it.", file.FullName)
					Throw cex

				Finally

					If Not metadataEditor Is Nothing Then

						metadataEditor.Flush()
						metadataEditor.Close()

					End If

				End Try

				Return True

			End Function

			Private Function InterrogateAttributes( _
				ByVal file As FileInfo, _
				Optional ByVal stream As UInt16 = UInt16.MaxValue _
			) As FileMetaData

				Dim return_Object As FileMetaData = Nothing
				Dim metadataEditor As IWMMetadataEditor2 = Nothing

				Dim attributes As New List(Of DictionaryEntry)

				attributes.Add(New DictionaryEntry(FileMetaData.PROPERTY_FILE, file))

				Try

					MediaCommands.CreateEditor(metadataEditor)

					metadataEditor.Open(file.FullName)

					' ---- Attributes ----
					Dim attribute_Count As UInt16
					CType(metadataEditor, IWMHeaderInfo3).GetAttributeCount(stream, attribute_Count)

					If attribute_Count > 0 Then

						For i As UInt16 = 0 To attribute_Count - 1

							Dim attribute_Type As WMT_ATTR_DATATYPE
							Dim attribute_Name As Byte() = Nothing
							Dim attribute_NameLength As UInt16
							Dim attribute_Value As Byte() = Nothing
							Dim attribute_ValueLength As UInt16

							CType(metadataEditor, IWMHeaderInfo3).GetAttributeByIndex( _
							i, stream, attribute_Name, attribute_NameLength, _
							attribute_Type, attribute_Value, attribute_ValueLength)

							attribute_Name = New Byte((attribute_NameLength * 2) - 1) {}
							attribute_Value = New Byte(attribute_ValueLength - 1) {}

							CType(metadataEditor, IWMHeaderInfo3).GetAttributeByIndex( _
							i, stream, attribute_Name, attribute_NameLength, _
							attribute_Type, attribute_Value, attribute_ValueLength)

							Dim actual_Name As String = GetString(attribute_Name).Replace(UNDER_SCORE, String.Empty).Replace(FORWARD_SLASH, String.Empty)
							Dim actual_Value As Object = GetValue(attribute_Value, attribute_Type)

							If Not actual_Value Is Nothing AndAlso Not (actual_Value.GetType Is GetType(String) AndAlso String.IsNullOrEmpty(actual_Value)) Then _
								attributes.Add(New DictionaryEntry(actual_Name, actual_Value))

						Next

					End If
					' --------------------

					' ---- Markers ----
					Dim marker_Count As UInt16
					CType(metadataEditor, IWMHeaderInfo3).GetMarkerCount(marker_Count)

					Dim markers As New List(Of FileMarker)

					If marker_Count > 0 Then

						For i As UInt16 = 0 To marker_Count - 1

							Dim marker_Name As Byte() = Nothing
							Dim marker_NameLength As UInt16 = 0
							Dim marker_Position As UInt64 = 0

							CType(metadataEditor, IWMHeaderInfo3).GetMarker(i, marker_Name, marker_NameLength, marker_Position)

							marker_Name = New Byte((marker_NameLength * 2) - 1) {}

							CType(metadataEditor, IWMHeaderInfo3).GetMarker(i, marker_Name, marker_NameLength, marker_Position)

							markers.Add(New FileMarker(GetString(marker_Name), marker_Position))		

						Next

					End If
					' -----------------

					' Create the return
					return_Object = TypeAnalyser.CreateAndPopulate(GetType(FileMetaData), attributes.ToArray)
					return_Object.Markers = markers

				Catch cex As COMException
				Finally

					If Not metadataEditor Is Nothing Then

						' metadataEditor.Flush()
						metadataEditor.Close()

					End If

				End Try

				Return return_Object

			End Function

		#End Region

		#Region " Command Processing Methods "

			<Command( _
				ResourceContainingType:=GetType(MediaCommands), _
				ResourceName:="CommandDetails", _
				Name:="get-metadata", _
				Description:="@commandMediaMetadataDescription@" _
			)> _
			Public Function ProcessCommandGetMetadata( _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFilePath@" _
				)> _
				ByVal mediaFile As IO.FileInfo _
			) As FileMetaData()

				Return New FileMetaData() {InterrogateAttributes(mediaFile)}

			End Function

			<Command( _
				ResourceContainingType:=GetType(MediaCommands), _
				ResourceName:="CommandDetails", _
				Name:="get-metadata", _
				Description:="@commandMediaMetadataDescription@" _
			)> _
			Public Function ProcessCommandGetMetadata( _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFileDirectory@" _
				)> _
				ByVal mediaDirectory As IO.DirectoryInfo, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFileSearchPattern@" _
				)> _
				ByVal mediaSearchPattern As String _
			) As FileMetaData()

				Return ProcessCommandGetMetadata(mediaDirectory, mediaSearchPattern, True)

			End Function

			<Command( _
				ResourceContainingType:=GetType(MediaCommands), _
				ResourceName:="CommandDetails", _
				Name:="get-metadata", _
				Description:="@commandMediaMetadataDescription@" _
			)> _
			Public Function ProcessCommandGetMetadata( _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFileDirectory@" _
				)> _
				ByVal mediaDirectory As IO.DirectoryInfo, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFileSearchPattern@" _
				)> _
				ByVal mediaSearchPattern As String, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFileSearchSubDirectories@" _
				)> _
				ByVal subDirectories As Boolean _
			) As FileMetaData()

				Dim files As IO.FileInfo()
				Dim returnList As New List(Of FileMetaData)

				files = mediaDirectory.GetFiles(mediaSearchPattern, SearchOption.TopDirectoryOnly)

				For i As Integer = 0 To files.Length - 1

				If files(i).Attributes Xor FileAttributes.System <> FileAttributes.System Then

					Dim interrogatedObject As FileMetaData = InterrogateAttributes(files(i))
					If Not interrogatedObject Is Nothing Then returnList.Add(interrogatedObject)

				End If

				Next

				If subDirectories Then

					Dim directories As IO.DirectoryInfo() = mediaDirectory.GetDirectories()

					For i As Integer = 0 To directories.Length - 1

						If (directories(i).Attributes And FileAttributes.System) <> FileAttributes.System Then

							returnList.AddRange(ProcessCommandGetMetadata(directories(i), mediaSearchPattern, True))

						End If

					Next

				End If

				Return returnList.ToArray

			End Function

			<Command( _
				ResourceContainingType:=GetType(MediaCommands), _
				ResourceName:="CommandDetails", _
				Name:="set-metadata", _
				Description:="@commandMediaSetMetadataDescription@" _
			)> _
			Public Function ProcessCommandSetMetadata( _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFilePath@" _
				)> _
				ByVal mediaFile As IO.FileInfo, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaAttributeName@" _
				)> _
				ByVal attribute As String, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaAttributeValue@" _
				)> _
				ByVal value As String _
			) As Boolean

				Return SetAttribute(mediaFile, attribute, value)

			End Function

			<Command( _
				ResourceContainingType:=GetType(MediaCommands), _
				ResourceName:="CommandDetails", _
				Name:="add-marker", _
				Description:="@commandMediaAddMarkerDescription@" _
			)> _
			Public Function ProcessCommandAddMarker( _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaFilePath@" _
				)> _
				ByVal mediaFile As IO.FileInfo, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaMarkerName@" _
				)> _
				ByVal name As String, _
				<Configurable( _
					ResourceContainingType:=GetType(MediaCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandMediaMarkerName@" _
				)> _
				ByVal location As TimeSpan _
			) As Boolean

				Dim existingFile As FileMetaData = InterrogateAttributes(mediaFile)

				For i As Integer = 0 To existingFile.Markers.Count - 1

					' Check them
					If String.Compare(existingFile.Markers(i).Name, name, True) = 0 AndAlso _
						existingFile.Markers(i).Location = location.Ticks Then Return False

				Next

				Dim metadataEditor As IWMMetadataEditor2 = Nothing

				Try

					MediaCommands.CreateEditor(metadataEditor)

					metadataEditor.Open(mediaFile.FullName)

					CType(metadataEditor, IWMHeaderInfo3).AddMarker( _
						GetBytes(name, WMT_ATTR_DATATYPE.WMT_TYPE_STRING), _
						Convert.ToUInt64(location.Ticks))

				Catch cex As COMException

					If (cex.ErrorCode = -1072889827) Then Throw New FileNotFoundException( _
						"Failed to open the file into memory. The file may  be missing or you may not have permission to open it.", mediaFile.FullName)

					Throw cex

				Finally

					If Not metadataEditor Is Nothing Then

						metadataEditor.Flush()
						metadataEditor.Close()

					End If

				End Try

				Return True

			End Function

		#End Region

		#Region " Private Shared Methods "

			Private Shared Function GetDataType( _
				ByVal prop As MemberAnalyser _
			) As WMT_ATTR_DATATYPE

				If prop.ReturnType Is GetType(System.UInt16) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_WORD

				ElseIf prop.ReturnType Is GetType(System.UInt32) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_DWORD

				ElseIf prop.ReturnType Is GetType(System.UInt64) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_QWORD

				ElseIf prop.ReturnType Is GetType(System.String) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_STRING

				ElseIf prop.ReturnType Is GetType(System.Boolean) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_BOOL

				ElseIf prop.ReturnType Is GetType(System.Guid) Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_GUID

				ElseIf prop.ReturnType Is GetType(System.Byte).MakeArrayType Then

					Return WMT_ATTR_DATATYPE.WMT_TYPE_BINARY

				Else

					Return WMT_ATTR_DATATYPE.WMT_TYPE_BINARY

				End If

			End Function

			Private Shared Function GetBytes( _
				ByVal value_String As String, _
				ByVal value_Type As WMT_ATTR_DATATYPE _
			) As Byte()

				Select Case value_Type

					Case WMT_ATTR_DATATYPE.WMT_TYPE_DWORD

						Dim return_Bytes(4) As Byte

						Buffer.BlockCopy(New UInt32() {Convert.ToUInt32(value_String)}, 0, return_Bytes, 0, return_Bytes.Length)

						Return return_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_STRING

						Dim return_Bytes((value_String.Length + 1) * 2) As Byte

						Buffer.BlockCopy(value_String.ToCharArray, 0, return_Bytes, 0, value_String.Length * 2)

						return_Bytes(return_Bytes.Length - 2) = 0
						return_Bytes(return_Bytes.Length - 1) = 0

						Return return_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_BOOL

						Dim return_Bytes(1) As Byte

						Buffer.BlockCopy(New Boolean() {Convert.ToBoolean(value_String)}, 0, return_Bytes, 0, return_Bytes.Length)

						Return return_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_QWORD

						Dim return_Bytes(8) As Byte

						Buffer.BlockCopy(New UInt64() {Convert.ToUInt64(value_String)}, 0, return_Bytes, 0, return_Bytes.Length)

						Return return_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_WORD

						Dim return_Bytes(2) As Byte

						Buffer.BlockCopy(New UInt16() {Convert.ToUInt16(value_String)}, 0, return_Bytes, 0, return_Bytes.Length)

						Return return_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_GUID

						Return New Guid(value_String).ToByteArray

					Case Else

						Return New Byte() {}

				End Select

			End Function

			Private Shared Function GetValue( _
				ByVal value_Bytes As Byte(), _
				ByVal value_Type As WMT_ATTR_DATATYPE _
			) As Object

				Select Case value_Type

					Case WMT_ATTR_DATATYPE.WMT_TYPE_DWORD

						Return BitConverter.ToUInt32(value_Bytes, 0)

					Case WMT_ATTR_DATATYPE.WMT_TYPE_STRING

						Return GetString(value_Bytes)

					Case WMT_ATTR_DATATYPE.WMT_TYPE_BINARY

						Return value_Bytes

					Case WMT_ATTR_DATATYPE.WMT_TYPE_BOOL

						Return BitConverter.ToBoolean(value_Bytes, 0)

					Case WMT_ATTR_DATATYPE.WMT_TYPE_QWORD

						Return BitConverter.ToUInt64(value_Bytes, 0)

					Case WMT_ATTR_DATATYPE.WMT_TYPE_WORD

						Return BitConverter.ToUInt16(value_Bytes, 0)

					Case WMT_ATTR_DATATYPE.WMT_TYPE_GUID

						Return New Guid(value_Bytes)

					Case Else

						Return Nothing

				End Select

			End Function

			Private Shared Function GetString( _
				ByVal string_Bytes As Byte() _
			) As String

				If Not string_Bytes Is Nothing AndAlso string_Bytes.Length >= 2 Then

					Dim stringLength As Integer

					' Deal with weird null/non-null terminated stuff.
					If string_Bytes(string_Bytes.Length - 2) = 0 AndAlso _
						string_Bytes(string_Bytes.Length - 1) = 0 Then

						stringLength = Math.Ceiling(string_Bytes.Length / 2) - 1

					Else

						stringLength = Math.Ceiling(string_Bytes.Length / 2)

					End If

					Dim return_Builder As New StringBuilder(stringLength)

					For i As Integer = 0 To stringLength - 1

						return_Builder.Append(BitConverter.ToChar(string_Bytes, i * 2))

					Next

					Return return_Builder.ToString

				End If

				Return Nothing

			End Function

		#End Region

		#Region " Public Shared Methods "

			<DllImport("WMVCore.dll", EntryPoint:="WMCreateEditor", CharSet:=CharSet.Unicode, _
				SetLastError:=True, ExactSpelling:=True, PreserveSig:=False)> _
			Public Shared Sub CreateEditor( _
				<Out(), MarshalAs(UnmanagedType.Interface)> ByRef ppEditor As IWMMetadataEditor2 _
			)
			End Sub

		#End Region

	End Class

End Namespace