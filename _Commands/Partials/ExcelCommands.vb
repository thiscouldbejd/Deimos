Imports Deimos.Excel
Imports Deimos.Excel.WorkbookWrapper
Imports Leviathan.Visualisation

Namespace Commands

	Partial Public Class ExcelCommands
		Implements IDisposable

		#Region " Public Command Methods "

			<Command( _
				ResourceContainingType:=GetType(ExcelCommands), _
				ResourceName:="CommandDetails", _
				Name:="input", _
				Description:="@commandExcelDescriptionInput@" _
			)> _
			Public Function ProcessCommandInput() As ICollection

				Return Worksheet.GetData(OriginX, OriginY, ExtentX, ExtentY, Nothing, Host)

			End Function

			<Command( _
				ResourceContainingType:=GetType(ExcelCommands), _
				ResourceName:="CommandDetails", _
				Name:="output", _
				Description:="@commandExcelDescriptionOutput@" _
			)> _
			Public Sub ProcessCommandOutput( _
				<Configurable( _
					ResourceContainingType:=GetType(ExcelCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandExcelParameterDescriptionFormattedObjects@" _
				)> _
				ByVal value As Cube _
			)

				If Not value Is Nothing Then Worksheet.OutputCube(New Cube() {value}, Host)

			End Sub

			<Command( _
				ResourceContainingType:=GetType(ExcelCommands), _
				ResourceName:="CommandDetails", _
				Name:="aggregate", _
				Description:="@commandExcelDescriptionAggregate@" _
			)> _
			Public Function ProcessCommandAggregate( _
				<Configurable( _
					ResourceContainingType:=GetType(ExcelCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandExcelParameterDescriptionDirectory@" _
				)> _
				ByVal inputDirectory As IO.DirectoryInfo _
			) As IList

				Dim input_Files As New List(Of IO.FileInfo)

				For i As Integer = 0 To EXCEL_SUFFIXES.Length - 1

					input_Files.AddRange(inputDirectory.GetFiles(String.Format("*.{0}", EXCEL_SUFFIXES(i))))

				Next

				Dim return_Type As Type = Nothing
				Dim return_List As New ArrayList()

				For i As Integer = 0 To input_Files.Count - 1

					Dim input_Book As WorkbookWrapper = Nothing

					If TryParse(input_Files(i).FullName, input_Book) Then

						Dim input_Sheets As WorksheetWrapper() = input_Book.Sheets

						For j As Integer = 0 To input_Sheets.Length - 1

							Dim return_Values As ICollection = Nothing
							return_Values = input_Sheets(j).GetData(OriginX, OriginY, ExtentX, ExtentY, return_Type, Host)

							If Not return_Values Is Nothing Then

								If return_Type Is Nothing AndAlso Not return_Values Is Nothing Then

									For Each single_Object As Object In return_Values

										return_Type = single_Object.GetType()
										Exit For

									Next

								End If

								return_List.AddRange(return_Values)

							End If

						Next

						input_Book.Dispose()
						input_Book = Nothing

					End If

				Next

				Return return_List

			End Function

		#End Region

	End Class

End Namespace