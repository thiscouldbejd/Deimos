Imports Leviathan.Commands.FilesCommands
Imports System.IO
Imports System.Runtime.InteropServices

Partial Public Class DocumentBase
	Implements IDisposable

	#Region " Public Shared Variables "

		''' <summary>
		''' Provides Access to the COM Identifier Class for Microsoft Excel.
		''' </summary>
		''' <remarks></remarks>
		Public Shared EXCEL_COM_IDENTIFIER As String = "Excel.Application"

		''' <summary>
		''' Provides Access to the Suffixes of Excel Files.
		''' </summary>
		''' <remarks></remarks>
		Public Shared EXCEL_SUFFIXES As String() = New String() {"xls", "xlsx"}

		''' <summary>
		''' Provides Access to the COM Identifier Class for Microsoft Word.
		''' </summary>
		''' <remarks></remarks>
		Public Shared WORD_COM_IDENTIFIER As String = "Word.Application"

		''' <summary>
		''' Provides Access to the Suffixes of Word Files.
		''' </summary>
		''' <remarks></remarks>
		Public Shared WORD_SUFFIXES As String() = New String() {"doc", "docx"}

		''' <summary>
		''' Provides Access to the COM Identifier Class for Microsoft Visio.
		''' </summary>
		''' <remarks></remarks>
		Public Shared VISIO_COM_IDENTIFIER As String = "Visio.Application"

		''' <summary>
		''' Provides Access to the Suffixes of Visio Files.
		''' </summary>
		''' <remarks></remarks>
		Public Shared VISIO_SUFFIXES As String() = New String() {"vsd", "vdx"}

		''' <summary>
		''' Provides Access to the COM Identifier Class for Microsoft Outlook.
		''' </summary>
		''' <remarks></remarks>
		Public Shared OUTLOOK_COM_IDENTIFIER As String = "Outlook.Application"

		''' <summary>
		''' Provides Access to the Timeout Wait Period.
		''' </summary>
		''' <remarks></remarks>
		Public Shared TIMEOUT_WAIT As Integer = 200

		''' <summary>
		''' Provides Access to the Timeout Retries Count.
		''' </summary>
		''' <remarks></remarks>
		Public Shared TIMEOUT_RETRIES As Integer = 2

	#End Region

	#Region " Protected MustOverride Methods "

		Protected MustOverride Sub SetDocument( _
			ByVal document As Object _
		)

	#End Region

	#Region " Protected Methods "

		''' <summary>
		''' Method to Populate the Wrapper with a Document.
		''' </summary>
		''' <param name="wrappedDocument">The Document being Wrapped.</param>
		''' <param name="documentState">The state of the Document.</param>
		''' <param name="ownsApplication">Whether this Document 'owns' the Application (e.g. created it and therefore will kill it).</param>
		''' <remarks></remarks>
		Protected Sub Populate( _
			ByVal wrappedDocument As Object, _
			ByVal documentState As OfficeDocumentState, _
			ByVal ownsApplication As Boolean, _
			ByVal isVisible As Boolean _
		)

			m_OwnsApplicationInstance = ownsApplication
			m_State = documentState
			m_ApplicationVisible = isVisible

			SetDocument(wrappedDocument)

		End Sub

		''' <summary>
		''' Method to Get an Office Document.
		''' </summary>
		''' <param name="documentName">The Name/Path of the Document.</param>
		''' <param name="applicationType">The Type of Application Document to get.</param>
		''' <remarks></remarks>
		Protected Friend Function ParseFromString( _
			ByVal documentName As String, _
			ByVal applicationType As OfficeApplication _
		) As DocumentBase

			' This determines whether we (as in this code) have created the application instance.
			Dim isNewInstance As Boolean
			Dim isVisible As Boolean

			' If the document name is suffixed with a '+' then we can create a file, otherwise not.
			Dim canCreateFile As Boolean = Not String.IsNullOrEmpty(documentName) AndAlso _
			(documentName.EndsWith(PLUS) OrElse documentName.EndsWith(HYPHEN))

			Dim canDeleteFile As Boolean = Not String.IsNullOrEmpty(documentName) AndAlso documentName.EndsWith(HYPHEN)
			If Not String.IsNullOrEmpty(documentName) Then documentName = documentName.TrimEnd(PLUS, HYPHEN)

			' Delete the existing file if possible
			If canDeleteFile AndAlso IO.File.Exists(documentName) Then IO.File.Delete(documentName)

			' Holds a Reference to the actual Application Object.
			' Create the Application Instance (if a new instance is created, that will be passed back
			' using the byref parameter). Not need to Null test as this is done in the
			' GetApplicationInstance routine.
			Dim officeApplication As Object = GetApplicationInstance(applicationType, isNewInstance, isVisible, False)

			''''''''''
			' Case 1 '
			''''''''''
			' No Document Name Supplied, so just created a new one.
			If State = D.None AndAlso String.IsNullOrEmpty(documentName) Then

				' No Name supplied, so we need to create a new file!
				Select Case applicationType

					Case A.Excel

						Populate(CType(officeApplication, E.Application).Workbooks.Add(), D.Created, isNewInstance, isVisible)

					Case A.Word

						Populate(CType(officeApplication, W.Application).Documents.Add(Visible:=True), D.Created, isNewInstance, isVisible)

					Case A.Visio

						Populate(CType(officeApplication, V.Application).Documents.Add(Nothing), D.Created, isNewInstance, isVisible)

				End Select

			End If

			If State = D.None Then

				' Get the potential path for the Document.
				Dim documentPath As String = GetFilePath(documentName)

				''''''''''
				' Case 2 '
				''''''''''
				' The Application was already open, so lets check all open documents.
				If Not isNewInstance Then

					' --- Check Existing Open Documents ---
					Dim documentCount As Integer = 0

					Select Case applicationType

						Case A.Excel

							documentCount = CType(officeApplication, E.Application).Workbooks.Count

						Case A.Word

							documentCount = CType(officeApplication, W.Application).Documents.Count

						Case A.Visio

							documentCount = CType(officeApplication, V.Application).Documents.Count

					End Select

					' Iterate through all existing Documents
					For i As Integer = 1 To documentCount

						' Firstly, get the full name/path of each document to compare.
						Dim currentName As String = Nothing
						Dim currentFullName As String = Nothing
						Dim currentDocument As Object = Nothing

						Select Case applicationType

							Case A.Excel

								currentDocument = CType(officeApplication, E.Application).Workbooks(i)
								currentName = CType(currentDocument, E.Workbook).Name
								currentFullName = CType(currentDocument, E.Workbook).FullName

							Case A.Word

								currentDocument = CType(officeApplication, W.Application).Documents(i)
								currentName = CType(currentDocument, W.Document).Name
								currentFullName = CType(currentDocument, W.Document).FullName

							Case A.Visio

								currentDocument = CType(officeApplication, V.Application).Documents(i)
								currentName = CType(currentDocument, V.Document).Name
								currentFullName = CType(currentDocument, V.Document).FullName

						End Select

						' Compare the paths and if they match, assign the correct document and exit the loop.
						If String.Compare(currentFullName, documentPath, True) = 0 OrElse _
							String.Compare(currentName, documentName, True) = 0 Then

							Populate(currentDocument, D.Existing, isNewInstance, isVisible)

							Exit For

						End If

					Next

				End If

				''''''''''
				' Case 3 '
				''''''''''
				' We'll try to load the file.
				If State = OfficeDocumentState.None AndAlso _
				(applicationType = A.Excel AndAlso CompareExtensions(documentName, EXCEL_SUFFIXES)) OrElse _
				(applicationType = A.Word AndAlso CompareExtensions(documentName, WORD_SUFFIXES)) OrElse _
				(applicationType = A.Visio AndAlso CompareExtensions(documentName, VISIO_SUFFIXES)) Then

					' If we're creating new, then see if we can open a new file.
					If File.Exists(documentPath) Then

						' If a file exists, then open it.
						Select Case applicationType

							Case A.Excel

								Populate(CType(officeApplication, E.Application).Workbooks.Open( _
									documentPath, ReadOnly:=(Not canCreateFile)), OfficeDocumentState.Opened, _
									isNewInstance, isVisible)

							Case A.Word

								Populate(CType(officeApplication, W.Application).Documents.Open( _
									documentPath, ReadOnly:=(Not canCreateFile)), OfficeDocumentState.Opened, _
									isNewInstance, isVisible)

							Case A.Visio

								Populate(CType(officeApplication, V.Application).Documents.Open( _
									documentPath), OfficeDocumentState.Opened, isNewInstance, isVisible)

						End Select

					ElseIf canCreateFile Then

						Dim currentDocument As Object = Nothing

						' The file doesn't exist in the location, so we should try to create it.
						Select Case applicationType

							Case A.Excel

								' Add the Document
								currentDocument = CType(officeApplication, E.Application).Workbooks.Add()

								' Save out the File.
								CType(currentDocument, E.Workbook).SaveAs(documentPath)

							Case A.Word

								' Add the Document
								currentDocument = CType(officeApplication, W.Application). _
								Documents.Add(Visible:=True)

								' Save out the File.
								CType(currentDocument, W.Document).SaveAs(documentPath)

							Case A.Visio

								' Add the Document
								currentDocument = CType(officeApplication, V.Application). _
								Documents.Add(Nothing)

								' Save out the File.
								CType(currentDocument, V.Document).SaveAs(documentPath)

						End Select

						' Create the File by Addition to the Collection.
						If Not currentDocument Is Nothing Then _
						Populate(currentDocument, OfficeDocumentState.Created, _
						isNewInstance, isVisible)

					End If

				End If

			End If

			' If the document is still nothing, throw a relevant exception.
			If State = D.None Then _
				Throw New Exception(String.Format(SingleResource( _
				GetType(DocumentBase), RESOURCEMANAGER_NAME_EXCEPTION, _
				"commandOfficeCouldNotFindFile"), documentName))

			Return Me

		End Function

	#End Region

	#Region " Public Shared Methods "

		''' <summary>
		''' Method to Get an Application Instance, either by getting an existing one
		''' or creating a new one.
		''' </summary>
		''' <param name="applicationType">The Type of the Application.</param>
		''' <param name="applicationCreated">Whether or not it was created (so it can be closed).</param>
		''' <param name="applicationVisible">Whether is should be visible.</param>
		''' <returns></returns>
		''' <remarks>Will throw an exception if required.</remarks>
		Public Shared Function GetApplicationInstance( _
			ByVal applicationType As OfficeApplication, _
			ByRef applicationCreated As Boolean, _
			ByRef applicationIsVisible As Boolean, _
			Optional ByVal applicationVisible As Boolean = True _
		) As Object

			Dim applicationObject As Object = Nothing
			Dim applicationClass As String = Nothing

			' This determines the Type of Application Identifier to use.
			Select Case applicationType

				Case A.Excel

					applicationClass = EXCEL_COM_IDENTIFIER

				Case A.Word

					applicationClass = WORD_COM_IDENTIFIER

				Case A.Visio

					applicationClass = VISIO_COM_IDENTIFIER

			End Select

			Try

				Dim currentTry As Integer = 0

				Do Until currentTry >= TIMEOUT_RETRIES

					Try

						' First, try to get an instance of the application.
						applicationObject = GetObject(Class:=applicationClass)

						' Favour NOT using NON-Visible instances as they're probably doing other stuff and
						' will be hard to kill/handle.
						If Not applicationObject Is Nothing Then

							Dim killApplication As Boolean

							Select Case applicationType

								Case A.Excel

									killApplication = Not CType(applicationObject, E.Application).Visible

								Case A.Word

									killApplication = Not CType(applicationObject, W.Application).Visible

								Case A.Visio

									killApplication = Not CType(applicationObject, V.Application).Visible

							End Select

							currentTry = TIMEOUT_RETRIES

							If killApplication Then

								Marshal.ReleaseComObject(applicationObject)
								applicationObject = Nothing

							End If

						End If

					Catch cex As InvalidCastException

						Threading.Thread.Sleep(TIMEOUT_WAIT)
						currentTry += 1

						If Not applicationObject Is Nothing Then
							Marshal.ReleaseComObject(applicationObject)
							applicationObject = Nothing
						End If

					End Try

				Loop

			Catch ex As Exception
			End Try

			If applicationObject Is Nothing Then

				Try

					' If the application is not present, create the application.
					applicationObject = CreateObject(applicationClass)
					applicationCreated = True
					applicationObject.Visible = applicationVisible

				Catch ex As Exception
				End Try

			End If

			If applicationObject Is Nothing Then _
				Throw New Exception( _
					String.Format(SingleResource( _
					GetType(DocumentBase), RESOURCEMANAGER_NAME_EXCEPTION, _
					"commandOfficeCouldNotGetApplicationInstance"), applicationClass))

			Select Case applicationType

				Case A.Excel

					applicationIsVisible = CType(applicationObject, E.Application).Visible

				Case A.Word

					applicationIsVisible = CType(applicationObject, W.Application).Visible

				Case A.Visio

					applicationIsVisible = Not CType(applicationObject, V.Application).Visible

			End Select

			Return applicationObject

		End Function

	#End Region

	#Region " IDisposable Implementation "

		Protected MustOverride Overloads Sub Dispose( _
			ByVal disposing As Boolean _
		)

			' This code added by Visual Basic to correctly implement the disposable pattern.
			Public Overloads Sub Dispose() Implements IDisposable.Dispose

			' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
			Dispose(True)
			GC.SuppressFinalize(Me)

		End Sub

	#End Region

End Class