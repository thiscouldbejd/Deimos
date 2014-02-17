Imports System.IO.File
Imports System.Runtime.InteropServices

Namespace Visio

	Partial Public Class DrawingWrapper

		#Region " Protected Properties "

			Protected ReadOnly Property FullName() As String
				Get
					If Not Drawing Is Nothing Then Return Drawing.FullName Else Return Nothing
				End Get
			End Property

		#End Region

		#Region " Public Properties "

			Public ReadOnly Property PageNames() As String()
				Get
					Dim l_Pages As V.Pages = Drawing.Pages

					Dim aryNames As String() = _
						Array.CreateInstance(GetType(String), l_Pages.Count)

					For i As Integer = 1 To aryNames.Length

						Dim l_Page As V.Page = l_Pages(i)

						aryNames(i - 1) = l_Page.Name

						l_Page = Nothing

					Next

					l_Pages = Nothing

					Return aryNames
				End Get
			End Property

		#End Region

		#Region " Protected Methods "

			''' <summary>
			''' Method to Set the Drawing (overrides base class).
			''' </summary>
			''' <param name="document">The drawing to set.</param>
			''' <remarks></remarks>
			Protected Overrides Sub SetDocument( _
				ByVal document As Object _
			)

				m_Drawing = document

				Dim l_Pages As V.Pages = Drawing.Pages
				m_PageCount = l_Pages.Count
				l_Pages = Nothing

			End Sub

		#End Region

		#Region " Public Methods "

			Public Sub Save()

				If Not String.IsNullOrEmpty(Drawing.Path) Then Drawing.Save()

			End Sub

			Public Sub Save( _
			ByVal filePath As String _
			)

				If Not String.IsNullOrEmpty(filePath) AndAlso Not Exists(filePath) Then _
				Drawing.SaveAs(filePath)

			End Sub

			Public Sub ClearPages( _
			Optional ByVal startIndex As Integer = 1 _
			)

				If startIndex <= PageCount Then

					Dim l_Pages As V.Pages = Drawing.Pages

					For i As Integer = startIndex To PageCount

						l_Pages.Delete(startIndex)

					Next

					l_Pages = Nothing

				End If

			End Sub

		#End Region

		#Region " Public Shared Methods "

			Public Shared Function TryParse( _
				ByVal value As String, _
				ByRef result As DrawingWrapper _
			) As Boolean

				Try

					result = New DrawingWrapper() _
						.ParseFromString(value, A.Visio)

					Return True

				Catch ex As Exception
					If Not result Is Nothing Then result.Dispose()
				Finally
				End Try

				Return False

			End Function

		#End Region

		#Region " IDisposable Implementation "

			Private disposedValue As Boolean = False        ' To detect redundant calls

			' IDisposable
			Protected Overloads Overrides Sub Dispose( _
				ByVal disposing As Boolean _
			)

				If Not disposedValue Then

					If Not m_Drawing Is Nothing Then

						If State <> OfficeDocumentState.Existing Then

							' Get a reference to the Application
							Dim l_Application As V.Application = m_Drawing.Application

							m_Drawing.Close()

							If OwnsApplicationInstance Then

								Dim l_Documents As V.Documents = l_Application.Documents

								If l_Documents.Count = 0 Then

									' Kill Documents Reference
									Marshal.ReleaseComObject(l_Documents)
									l_Documents = Nothing

									' Close Application
									l_Application.DisplayAlerts = False
									l_Application.Quit()

								End If

							End If

							' Kill Application Reference.
							Marshal.ReleaseComObject(l_Application)
							l_Application = Nothing

						End If

						' Kill Drawing Reference
						Marshal.ReleaseComObject(m_Drawing)
						m_Drawing = Nothing

					End If

				End If

				disposedValue = True

			End Sub

		#End Region

	End Class

End Namespace