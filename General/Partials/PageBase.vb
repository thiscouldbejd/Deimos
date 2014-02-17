Imports Deimos.Excel
Imports Deimos.Visio
Imports Hermes.Cryptography.Cipher
Imports System.CodeDom
Imports System.CodeDom.Compiler.CodeDomProvider
Imports System.IO

Partial Public Class PageBase
	Implements IDisposable

	#Region " Private Constants "

		Private Const SCHEMA_COLUMN_COLUMNNAME As String = "ColumnName"

		Private Const SCHEMA_COLUMN_COLUMNSIZE As String = "ColumnSize"

		Private Const SCHEMA_COLUMN_DATATYPE As String = "DataType"

		Private Const SCHEMA_COLUMN_BASEDATATYPE As String = "ProviderSpecificDataType"

		Private Const SCHEMA_COLUMN_NULLABLE As String = "AllowDBNull"

	#End Region

	#Region " Protected MustOverride Methods "

		Protected MustOverride Sub SetPage( _
			ByVal page As Object _
		)

	#End Region

	#Region " Protected Methods "

		''' <summary>
		''' Method to Populate the Wrapper with a Page.
		''' </summary>
		''' <param name="wrappedPage">The Page being Wrapped.</param>
		''' <param name="pageState">The state of the Page.</param>
		''' <remarks></remarks>
		Protected Sub Populate( _
			ByVal wrappedPage As Object, _
			ByVal pageState As OfficePageState, _
			ByVal Optional pageName As String = Nothing _
		)

			m_State = pageState

			If Not String.IsNullOrEmpty(pageName) Then wrappedPage.Name = pageName

			SetPage(wrappedPage)

		End Sub

		''' <summary>
		''' Method to Populate the Wrapper with a Page.
		''' </summary>
		''' <param name="parent">The Parent of the Page being Wrapped.</param>
		''' <param name="wrappedPage">The Page being Wrapped.</param>
		''' <param name="pageState">The state of the Page.</param>
		''' <remarks></remarks>
		Protected Sub Populate( _
			ByVal parent As Object, _
			ByVal wrappedPage As Object, _
			ByVal pageState As OfficePageState, _
			ByVal Optional pageName As String = Nothing _
		)

			m_Parent = parent

			m_State = pageState

			If Not String.IsNullOrEmpty(pageName) Then wrappedPage.Name = pageName

			SetPage(wrappedPage)

		End Sub

		''' <summary>
		''' Method to Get an Office Page.
		''' </summary>
		''' <param name="pageName">The Name of the Page.</param>
		''' <param name="applicationType">The Type of Application Page to get.</param>
		''' <remarks></remarks>
		Protected Function ParseFromString( _
			ByVal pageName As String, _
			ByVal containingDocument As DocumentBase, _
			ByVal applicationType As OfficeApplication _
		) As PageBase

			' If the document name is suffixed with a '+' then we can create a file, otherwise not.
			Dim canCreatePage As Boolean = Not String.IsNullOrEmpty(pageName) AndAlso pageName.EndsWith(PLUS)
			If Not String.IsNullOrEmpty(pageName) Then pageName = pageName.TrimEnd(PLUS)

			''''''''''
			' Case 1 '
			''''''''''
			' No Page Name Supplied, so just created a new one.
			If String.IsNullOrEmpty(pageName) OrElse containingDocument.State = D.Created Then

				Select Case applicationType

					Case A.Excel

						If containingDocument.State = D.Created Then CType(containingDocument, WorkbookWrapper).ClearSheets(2)
						Populate(CType(containingDocument, WorkbookWrapper).Workbook.Sheets.Item(1), P.Created, pageName)

					Case A.Visio

						If containingDocument.State = D.Created Then CType(containingDocument, DrawingWrapper).ClearPages(2)
						Populate(CType(containingDocument, DrawingWrapper).Drawing.Pages.Item(1), P.Created, pageName)

				End Select

			End If

			''''''''''
			' Case 2 '
			''''''''''
			' Search for existing page.
			If State = P.None Then

				' --- Check Existing Pages ---
				Dim pageCount As Integer = 0

				Select Case applicationType

					Case A.Excel

						pageCount = CType(containingDocument, WorkbookWrapper).Workbook.Sheets.Count

					Case A.Visio

						pageCount = CType(containingDocument, DrawingWrapper).Drawing.Pages.Count

				End Select

				' Iterate through all existing Pages
				For i As Integer = 1 To pageCount

					' Firstly, get the name of each page to compare.
					Dim currentName As String = Nothing
					Dim currentPage As Object = Nothing

					Select Case applicationType

						Case A.Excel

							currentPage = CType(containingDocument, WorkbookWrapper).Workbook.Sheets(i)
							currentName = CType(currentPage, E.Worksheet).Name

						Case A.Visio

							currentPage = CType(containingDocument, DrawingWrapper).Drawing.Pages(i)
							currentName = CType(currentPage, V.Page).Name

					End Select

					' Compare the names and if they match, assign the correct page and exit the loop.
					If String.Compare(currentName, pageName, True) = 0 Then

						Populate(currentPage, P.Existing)

						Exit For

					End If

				Next
				' -------------------------------------

			End If

			''''''''''
			' Case 3 '
			''''''''''
			' Search for existing page.
			If State = P.None AndAlso canCreatePage Then

				Select Case applicationType

					Case A.Excel

						Populate(CType(containingDocument, WorkbookWrapper).Workbook.Sheets.Add(After:=CType(containingDocument, WorkbookWrapper). _
							WorkBook.ActiveSheet), P.Created, pageName)

					Case A.Visio

						Populate(CType(containingDocument, DrawingWrapper).Drawing.Pages.Add(), P.Created, pageName)

				End Select

			End If

			' If the page is still nothing, throw an exception.
			If State = P.None Then Throw New Exception(String.Format(SingleResource(GetType(DocumentBase), RESOURCEMANAGER_NAME_EXCEPTION, _
				"commandOfficeCouldNotFindPage"), pageName))

			Return Me

		End Function

	#End Region

	#Region " Protected Shared Methods "

		Protected Shared Function GetObjectsFromDataTable( _
			ByVal dt As DataTable, _
			ByVal type As System.Type, _
			Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
		) As Array

			Dim returnTypeAnalyser As TypeAnalyser = TypeAnalyser.GetInstance(type)

			Dim parser As New FromString()

			Dim writeableColumns(dt.Columns.Count - 1) As MemberAnalyser

			For i As Integer = 0 To dt.Columns.Count - 1

				writeableColumns(i) = returnTypeAnalyser.GetMember(MemberAnalyser.SafeMemberName(dt.Columns(i).ColumnName))

			Next

			Dim returnArray As Array = Array.CreateInstance(type, dt.Rows.Count)

			Dim rowCount As Integer = dt.Rows.Count

			For i As Integer = 0 To rowCount - 1

				returnArray(i) = returnTypeAnalyser.Create

				For j As Integer = 0 To writeableColumns.Length - 1

					Dim value As Object = dt.Rows(i).Item(j)

					If Not value Is Nothing AndAlso Not IsDBNull(value) Then

						If value.GetType Is GetType(String) AndAlso Not writeableColumns(j).ReturnType Is GetType(String) Then

							Dim success As Boolean

							value = parser.Parse(value, success, writeableColumns(j).ReturnType)

							If success Then writeableColumns(j).Write(returnArray(i), value)

						Else

							writeableColumns(j).Write(returnArray(i), value)

						End If

					End If

					If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive) Then Host.Progress(i + 1 / rowCount, _
						"Populating Object Collection")

				Next

			Next

			Return returnArray

		End Function

		Protected Shared Function GetObjectsFromDataTable( _
			ByVal dt As DataTable, _
			Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
		) As Array

			Return GetObjectsFromDataTable(dt, GetTypeFromDataTable(dt), Host)

		End Function

		Protected Shared Function GetTypeFromDataTable( _
			ByVal dt As DataTable _
		) As Type

			Dim typeName As String = dt.TableName

			If String.IsNullOrEmpty(typeName) Then

				typeName = "generated_" & Create_Password(5, 0).ToLower

			Else

				typeName = typeName.Replace(SPACE, UNDER_SCORE)

			End If

			Dim generated_Class As New CodeTypeDeclaration(typeName)
			generated_Class.IsClass = True

			If dt.Columns.Count >= 6 AndAlso dt.Columns(0).ColumnName = SCHEMA_COLUMN_COLUMNNAME AndAlso dt.Columns(2).ColumnName = _
				SCHEMA_COLUMN_COLUMNSIZE AndAlso dt.Columns(6).ColumnName = SCHEMA_COLUMN_DATATYPE Then

				' This is a schema table, probably derived from an IDataReader
				For i As Integer = 0 To dt.Rows.Count - 1

					Dim field_Type As Type = dt.Rows(i)(SCHEMA_COLUMN_DATATYPE)
					Dim field_Name As String = MemberAnalyser.SafeMemberName(dt.Rows(i)(SCHEMA_COLUMN_COLUMNNAME).ToString)

					Dim generated_Field As New CodeMemberField(field_Type, field_Name)
					generated_Field.Attributes = MemberAttributes.Public
					generated_Class.Members.Add(generated_Field)

				Next

			Else

				For i As Integer = 0 To dt.Columns.Count - 1

					Dim field_Type As Type = dt.Columns(i).DataType
					Dim field_Name As String = MemberAnalyser.SafeMemberName(dt.Columns(i).ColumnName)

					Dim generated_Field As New CodeMemberField(field_Type, field_Name)
					generated_Field.Attributes = MemberAttributes.Public
					generated_Class.Members.Add(generated_Field)

				Next

			End If

			Dim generated_Unit As New CodeCompileUnit()
			Dim generated_Namespace As New CodeNamespace()
			generated_Namespace.Types.Add(generated_Class)
			generated_Unit.Namespaces.Add(generated_Namespace)

			Dim generated_Params As New Compiler.CompilerParameters()
			generated_Params.GenerateInMemory = True

			Dim generated_Provider As Compiler.CodeDomProvider = CreateProvider(OutputLanguage.VisualBasic.ToString)

			Try

				Dim generation_Result As Compiler.CompilerResults = generated_Provider.CompileAssemblyFromDom(generated_Params, generated_Unit)

				If Not generation_Result.Errors.HasErrors Then Return generation_Result.CompiledAssembly.GetType(typeName)

			Catch ex As Exception
			End Try

			Return Nothing

		End Function

	#End Region

	#Region " Public Shared Methods "

		''' <summary>
		''' Method to Parse the Document and Page name.
		''' </summary>
		''' <param name="documentAndPageName">The combined name.</param>
		''' <param name="documentName">The ByRef (Out) Parameter for the Document Name.</param>
		''' <param name="pageName">The ByRef (Out) Parameter for the Page Name.</param>
		''' <remarks></remarks>
		Public Shared Sub ParseName( _
			ByVal documentAndPageName As String, _
			ByRef documentName As String, _
			ByRef pageName As String _
		)

			If Not String.IsNullOrEmpty(documentAndPageName) AndAlso documentAndPageName.Contains(EXCLAMATION_MARK) Then

				pageName = documentAndPageName.Split(EXCLAMATION_MARK)(1)
				documentName = documentAndPageName.Split(EXCLAMATION_MARK)(0)

			Else

				pageName = Nothing
				documentName = documentAndPageName

			End If

		End Sub

	#End Region

	#Region " IDisposable Implementation "

		Protected MustOverride Overloads Sub Dispose( _
			ByVal disposing As Boolean _
		)

		#Region " IDisposable Support "

			' This code added by Visual Basic to correctly implement the disposable pattern.
			Public Overloads Sub Dispose() Implements IDisposable.Dispose

				' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
				Dispose(True)
				GC.SuppressFinalize(Me)

			End Sub

		#End Region

	#End Region

End Class