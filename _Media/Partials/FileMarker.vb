Namespace Media

	Partial Public Class FileMarker

		#Region " Public Properties "

			Public ReadOnly Property DisplayLocation() As String
				Get
					Return New TimeSpanConvertor().ParseStringFromTimespan(New TimeSpan(Location), New Boolean)
				End Get
			End Property

		#End Region

	End Class

End Namespace