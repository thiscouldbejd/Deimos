Imports Leviathan.Visualisation
Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Threading.Thread

Namespace Excel

	Partial Public Class WorksheetWrapper
		Implements IDisposable

		#Region " Protected Constants "

			Protected Const COM_ERROR_REJECTED As String = "RPC_E_CALL_REJECTED"

		#End Region

		#Region " Protected Methods "

			Protected Overrides Sub SetPage( _
				ByVal page As Object _
			)

				m_Worksheet = page

			End Sub

			Protected Function GetStartRow( _
				ByVal columnCount As Integer, _
				ByVal rowCount As Integer, _
				ByRef range As Object(,) _
			) As Integer

				Dim highestValues As Integer
				Dim highestValuedRow As Integer

				For i As Integer = 1 To rowCount

					Dim rowValues As Integer = 0

					For j As Integer = 1 To columnCount

						If Not range(i, j) Is Nothing Then rowValues += 1

					Next

					If rowValues = columnCount Then

						highestValuedRow = i
						Exit For

					ElseIf rowValues > highestValues Then

						highestValues = rowValues
						highestValuedRow = i

					End If

				Next

				Return highestValuedRow

			End Function

			Protected Sub WriteTable( _
				ByVal value As Slice, _
				ByVal columns As List(Of FormatterProperty), _
				Optional ByVal startColumn As Integer = 0, _
				Optional ByVal startRow As Integer = 0, _
				Optional ByVal endColumn As Integer = 0, _
				Optional ByVal endRow As Integer = 0, _
				Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
			)

				Dim currentRange As E.Range = Worksheet.UsedRange
				Dim currentColumns As E.Range = currentRange.Columns
				Dim currentRows As E.Range = currentRange.Rows

				If startColumn <= 0 Then

					If currentRange.Column = 1 AndAlso currentRange.Row = 1 AndAlso currentRange.Value2 Is Nothing Then

						startColumn = currentRange.Column

					Else

						startColumn = currentRange.Column + currentColumns.Count

					End If

				End If

				If startRow <= 0 Then

					If currentRange.Column = 1 AndAlso currentRange.Row = 1 AndAlso currentRange.Value2 Is Nothing Then

						startRow = currentRange.Row

					Else

						startRow = currentRange.Row + currentRows.Count

					End If

				End If

				Dim currentX As Integer = startColumn
				Dim currentY As Integer = startRow

				For j As Integer = 0 To columns.Count - 1

					SetCallValue(currentY, currentX + j, ObjectToSingleString(columns(j).DisplayName, " | "), currentRange)

				Next

				currentX = startColumn
				currentY += 1

				Dim rowCount As Integer = value.Rows.Count

				For j As Integer = 0 To rowCount - 1

					If Not value.Rows(j) Is Nothing AndAlso value.Rows(j).Cells.Count > 0 Then

						For k As Integer = 0 To value.Rows(j).Cells.Count - 1

							SetCallValue(currentY + j, currentX + k, ObjectToSingleString(value.Rows(j)(k), "; "), currentRange)

						Next

					End If

					If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive) Then Host.Progress(j + 1 / rowCount, "Outputting Table Rows")

				Next

				For j As Integer = 0 To columns.Count - 1

					Dim fitRange As E.Range = Worksheet.Columns(currentX + j)

					fitRange.AutoFit()

					Marshal.ReleaseComObject(fitRange)
					fitRange = Nothing

				Next

				Marshal.ReleaseComObject(currentColumns)
				currentColumns = Nothing

				Marshal.ReleaseComObject(currentRows)
				currentRows = Nothing

				Marshal.ReleaseComObject(currentRange)
				currentRange = Nothing

				If Parent.State <> D.Existing Then CType(Parent, WorkbookWrapper).Save()

			End Sub

		#End Region

		#Region " Public Methods "

			Public Function GetCellValues( _
				ByRef range As E.Range _
			) As Object

				Dim retValue As Object = Nothing

				If Parent.OwnsApplicationInstance AndAlso Not Parent.ApplicationVisible Then

					retValue = range.Value

				Else ' Use Interactive Mode if we don't own the Application Instance (e.g. we created it)

					Dim startTime As DateTime = Now
					Dim completed As Boolean = False

					While Not completed AndAlso Now.Subtract(startTime).Seconds < Timeout

						Try

							retValue = range.Value
							completed = True

						Catch ex As System.Runtime.InteropServices.COMException

							If ex.ErrorCode = -2147418111 OrElse ex.Message.Contains(COM_ERROR_REJECTED) Then Sleep(500)

						End Try

					End While

				End If

				Return retValue

			End Function

			Public Sub SetCallValue( _
				ByVal rowIndex As Integer, _
				ByVal columnIndex As Integer, _
				ByVal value As Object, _
				ByRef range As E.Range _
			)

				If Not value Is Nothing AndAlso value.GetType Is GetType(String) AndAlso Not String.IsNullOrEmpty(value) AndAlso CStr(value).Chars(0) = _
					DIGIT_ZERO AndAlso IsNumeric(value) Then value = QUOTE_SINGLE & value

				If Parent.OwnsApplicationInstance AndAlso Not Parent.ApplicationVisible Then

					range(rowIndex, columnIndex).Value = value

				Else ' Use Interactive Mode if we don't own the Application Instance (e.g. we created it)

					Dim startTime As DateTime = Now
					Dim completed As Boolean = False

					While Not completed AndAlso Now.Subtract(startTime).Seconds < Timeout

						Try

							range(rowIndex, columnIndex).Value = value

							completed = True

						Catch ex As System.Runtime.InteropServices.COMException

							If ex.ErrorCode = -2147418111 OrElse ex.Message.Contains(COM_ERROR_REJECTED) Then System.Threading.Thread.Sleep(500)

						End Try

					End While

				End If

			End Sub

			Public Function GetData( _
				Optional ByVal startColumn As Integer = 0, _
				Optional ByVal startRow As Integer = 0, _
				Optional ByVal endColumn As Integer = 0, _
				Optional ByVal endRow As Integer = 0, _
				Optional ByVal returnType As System.Type = Nothing, _
				Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing, _
				Optional ByVal mappings As Dictionary(Of System.String, System.String) = Nothing _
			) As ICollection

				' -- Get the Range of Values we'll be dealing with --
				Dim cellRange As E.Range = Worksheet.UsedRange
				Dim dataRange(,) As Object = GetCellValues(cellRange)
				cellRange = Nothing
				
				Dim columnCount As Integer = 0
				Dim rowCount As Integer = 0
				
				If Not dataRange Is Nothing Then
					columnCount = dataRange.GetUpperBound(1)
					rowCount = dataRange.GetUpperBound(0)
				End If
				' ---------------------------------------------------

				' -- Reset the Start and End Values if Required --
				If startColumn <= 0 Then startColumn = 1
				If startRow <= 0 Then startRow = GetStartRow(columnCount, rowCount, dataRange)
				If endColumn <= 0 Then endColumn = columnCount
				If endRow <= 0 Then endRow = rowCount
				' ------------------------------------------------

				'-- Test for Table Headings --
				Dim hasColumnHeadings As Boolean

				If rowCount > 1 Then

					Dim allString As Boolean = True

					For i As Integer = 1 To columnCount

						If Not dataRange(startRow, i) Is Nothing Then

							If dataRange(startRow, i).GetType Is GetType(String) AndAlso Not IsNumeric(dataRange(startRow, i)) Then

								If Not dataRange(startRow + 1, i) Is Nothing Then

									If Not dataRange(startRow + 1, i).GetType Is GetType(String) OrElse IsNumeric(dataRange(startRow + 1, i)) Then

										hasColumnHeadings = True
										Exit For

									End If

								End If

							Else

								allString = False

							End If

						End If

					Next

					If Not hasColumnHeadings AndAlso allString Then hasColumnHeadings = True

				End If
				'-----------------------

				' Create a New DataTable
				Dim dt As New DataTable()

				If columnCount > 0 Then

					' Declare Current Row
					Dim cRow As Integer = startRow

					' Perform an until Column Finish Loop
					For i As Integer = startColumn To columnCount

						If hasColumnHeadings AndAlso Not String.IsNullOrEmpty(dataRange(cRow, i)) Then

							' Add a column with a Name
							dt.Columns.Add(New DataColumn(dataRange(cRow, i)))

						Else

							' Add a column with an Integer
							dt.Columns.Add(New DataColumn(i - startColumn))

						End If

					Next

					' If there are Column Names, we need to move on a row to get to the data
					If hasColumnHeadings Then cRow += 1

					For i As Integer = cRow To rowCount

						Dim aryRowData(columnCount - 1) As Object

						For j As Integer = startColumn To columnCount

							aryRowData(j - startColumn) = dataRange(i, j)

						Next

						' Create and Populate a new Row in the Table
						dt.Rows.Add(aryRowData)

						' Iterate the Row Number
						cRow += 1

						If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive)  Then

							Dim shouldOutput As Integer = 0
							Math.DivRem(cRow - startRow, CInt(Math.Ceiling(rowCount / 50)), shouldOutput)
							If shouldOutput = 0 OrElse (cRow - startRow) / (endRow - startRow) = 1 Then host.Progress((cRow - startRow) / (endRow - startRow), _
								"Extracting Worksheet Rows")

						End If

					Next

				End If

				If Not mappings Is Nothing AndAlso mappings.Count > 0 Then

					For i As Integer = 0 To dt.Columns.Count - 1

						If mappings.ContainsKey(dt.Columns(i).ColumnName) Then _
							dt.Columns(i).ColumnName = mappings(dt.Columns(i).ColumnName)

					Next

				End If

				If returnType Is Nothing Then

					Return GetObjectsFromDataTable(dt, host)

				Else

					Return GetObjectsFromDataTable(dt, returnType, host)

				End If

			End Function

			Public Sub OutputCube( _
				ByVal values As Cube(), _
				Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
			)

				For i As Integer = 0 To values.Length - 1

					WriteTable(values(i).LastSlice, values(i).Columns, 0, 0, 0, 0, host)

				Next

			End Sub

		#End Region

		#Region " Public Shared Methods "

			Public Shared Function TryParseAll( _
				ByVal value As String, _
				ByRef results As WorksheetWrapper() _
			) As Boolean

				Dim book As WorkbookWrapper = Nothing

				If WorkbookWrapper.TryParse(value, book) Then

					Try
						results = GetSheets(book)

						If Not results Is Nothing Then Return True

					Catch ex As Exception

						If Not book Is Nothing Then book.Dispose()

					Finally
					End Try

				End If

				Return False

			End Function

			Public Shared Function TryParse( _
				ByVal value As String, _
				ByRef result As WorksheetWrapper _
			) As Boolean

				Dim documentName As String = Nothing, pageName As String = Nothing

				PageBase.ParseName(value, documentName, pageName)

				Dim book As WorkbookWrapper = Nothing

				If WorkbookWrapper.TryParse(documentName, book) Then

					Try

						result = New WorksheetWrapper().ParseFromString(pageName, book, A.Excel)

						result.Parent = book

						Return True

					Catch ex As Exception

						If Not book Is Nothing Then book.Dispose()

					Finally
					End Try

				End If

				Return False

			End Function

			Public Shared Function GetSheets( _
				ByVal value As WorkbookWrapper _
			) As WorkSheetWrapper()

				Dim l_Sheets As E.Sheets = value.Workbook.Sheets

				Dim arySheets As WorksheetWrapper() = Array.CreateInstance(GetType(WorksheetWrapper), l_Sheets.Count)

				For i As Integer = 1 To arySheets.Length

					arySheets(i - 1) = New WorksheetWrapper()
					arySheets(i - 1).Populate(value, l_Sheets(i), OfficePageState.Existing, l_Sheets(i).Name)

				Next

				l_Sheets = Nothing

				Return arySheets

			End Function

		#End Region

		#Region " IDisposable Implementation "

			Private disposedValue As Boolean = False        ' To detect redundant calls

			' IDisposable
			Protected Overloads Overrides Sub Dispose( _
				ByVal disposing As Boolean _
			)

				If Not disposedValue Then

					If Not m_Worksheet Is Nothing Then

						Marshal.ReleaseComObject(m_Worksheet)
						m_Worksheet = Nothing

					End If

					If Not Parent Is Nothing Then Parent.Dispose()

				End If

				disposedValue = True

			End Sub

		#End Region

	End Class

End Namespace