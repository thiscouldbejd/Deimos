Imports System.IO.File
Imports System.Runtime.InteropServices

Namespace Excel

	Partial Public Class WorkbookWrapper

		#Region " Public Properties "

			Public ReadOnly Property FullName() As String
				Get
					If Not Workbook Is Nothing Then Return Workbook.FullName Else Return Nothing
				End Get
			End Property

			Public ReadOnly Property Sheets() As WorksheetWrapper()
				Get
					Return WorkSheetWrapper.GetSheets(Me)
				End Get
			End Property

			Public ReadOnly Property SheetNames() As String()
				Get
					Dim l_Sheets As E.Sheets = Workbook.Sheets

					Dim aryNames As String() = _
						Array.CreateInstance(GetType(String), l_Sheets.Count)

					For i As Integer = 1 To aryNames.Length

						Dim l_Sheet As E.Worksheet = l_Sheets(i)

						aryNames(i - 1) = l_Sheet.Name

						l_Sheet = Nothing

					Next

					l_Sheets = Nothing

					Return aryNames
				End Get
			End Property

		#End Region

		#Region " Protected Methods "

			''' <summary>
			''' Method to Set the Workbook (overrides base class).
			''' </summary>
			''' <param name="document">The workbook to set.</param>
			''' <remarks></remarks>
			Protected Overrides Sub SetDocument( _
				ByVal document As Object _
			)

				m_Workbook = document

				Dim l_Sheets As E.Sheets = Workbook.Sheets
				m_SheetCount = l_Sheets.Count
				l_Sheets = Nothing

			End Sub

		#End Region

		#Region " Public Methods "

			Public Sub Save()

				If Not String.IsNullOrEmpty(Workbook.Path) Then Workbook.Save()

			End Sub

			Public Sub Save( _
				ByVal filePath As String _
			)

				If Not String.IsNullOrEmpty(filePath) AndAlso Not Exists(filePath) Then _
					Workbook.SaveAs(filePath)

			End Sub

			Public Sub ClearSheets( _
				Optional ByVal startIndex As Integer = 1 _
			)

				If startIndex <= SheetCount Then

					Dim l_Sheets As E.Sheets = Workbook.Sheets

					For i As Integer = startIndex To SheetCount

						Dim l_Sheet As E.Worksheet = l_Sheets.Item(startIndex)
						l_Sheet.Delete
						l_Sheet = Nothing

					Next

					l_Sheets = Nothing

				End If

			End Sub

		#End Region

		#Region " Public Shared Methods "

			Public Overloads Shared Function TryParse( _
				ByVal value As String, _
				ByRef result As WorkbookWrapper _
			) As Boolean

				Try

					result = New WorkbookWrapper().ParseFromString(value, A.Excel)

					Return True

				Catch ex As Exception
					If Not result Is Nothing Then result.Dispose()
				Finally
				End Try

				Return False

			End Function

		#End Region

		#Region " IDisposable Implementation "

			Private disposedValue As Boolean = False ' To detect redundant calls

			' IDisposable
			Protected Overloads Overrides Sub Dispose( _
				ByVal disposing As Boolean _
			)

				If Not disposedValue Then

					If Not m_Workbook Is Nothing Then

						If State <> OfficeDocumentState.Existing Then

							' Get a reference to the Application
							Dim l_Application As E.Application = m_Workbook.Application

							m_Workbook.Close(SaveChanges:=False)

							If OwnsApplicationInstance Then

								Dim l_Workbooks As E.Workbooks = l_Application.Workbooks

								If l_Workbooks.Count = 0 Then

									' Kill Workbooks Reference
									Marshal.ReleaseComObject(l_Workbooks)
									l_Workbooks = Nothing

									' Close Application
									l_Application.DisplayAlerts = False
									l_Application.Quit()

								End If

							End If

							' Kill Application Reference.
							Marshal.ReleaseComObject(l_Application)
							l_Application = Nothing

						End If

						' Kill Workbook Reference
						Marshal.ReleaseComObject(m_Workbook)
						m_Workbook = Nothing

					End If

				End If

				disposedValue = True

			End Sub

		#End Region

	End Class

End Namespace