Namespace Media

	Partial Public Class FileMetaData

		#Region " Public Properties "

			Public ReadOnly Property Path() As String
				Get
					If Not File Is Nothing Then Return File.FullName Else Return Nothing
				End Get
			End Property

			Public ReadOnly Property Name() As String
				Get
					If Not File Is Nothing Then Return File.Name Else Return Nothing
				End Get
			End Property

			Public ReadOnly Property Size() As System.Int64
				Get
					If Not File Is Nothing Then Return File.Length Else Return 0
				End Get
			End Property

			Public ReadOnly Property DisplaySize() As String
				Get
					Return New LongConvertor().ParseStringFromLong(Math.Round(Size / 1024), New Boolean, Nothing) & "kb"
				End Get
			End Property

			Public ReadOnly Property DisplayDuration() As String
				Get
					Return New TimeSpanConvertor().ParseStringFromTimespan(New TimeSpan(Duration), New Boolean)
				End Get
			End Property

		#End Region

	End Class

End Namespace