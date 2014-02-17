Imports System.Reflection
Imports System.Xml

Namespace Excel

	Partial Public Class XmlDocument

		#Region " Public Shared Constants "

			''' <summary>
			''' Public Constant Defining the Office Schema Namespace.
			''' </summary>
			''' <remarks></remarks>
			Public Const OFFICE_NAMESPACE As String = "urn:schemas-microsoft-com:office:office"

			''' <summary>
			''' Public Constant Defining the Excel Schema Namespace.
			''' </summary>
			''' <remarks></remarks>
			Public Const EXCEL_NAMESPACE As String = "urn:schemas-microsoft-com:office:excel"

			''' <summary>
			''' Public Constant Defining the Spreadsheet Schema Namespace.
			''' </summary>
			''' <remarks></remarks>
			Public Const SPREADSHEET_NAMESPACE As String = "urn:schemas-microsoft-com:office:spreadsheet"

			''' <summary>
			''' Public Constant Defining the Html Schema Namespace.
			''' </summary>
			''' <remarks></remarks>
			Public Const HTML_NAMESPACE As String = "http://www.w3.org/TR/REC-html40"

		#End Region

	#Region " Private Properties "

		''' <summary>
		''' Provides Access to the Xml Writer which will do the writing.
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Private ReadOnly Property GetXmlWriter() As XmlTextWriter
			Get
				Return New System.Xml.XmlTextWriter(Stream, System.Text.Encoding.Unicode)
			End Get
		End Property

		''' <summary>
		''' Provides Access to a Boolean Value indicating whether a Visible Property List will be used.
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Private ReadOnly Property UsingPropertyList() As Boolean
			Get
				Return Not m_Properties.Count = 0
			End Get
		End Property

	#End Region

	#Region " Private Methods "

		''' <summary>
		''' Private Method to Create the Framework of the Excel Xml Document.
		''' </summary>
		''' <returns>An Xml Text Writer Wrapping the New Excel Xml Document.</returns>
		''' <remarks></remarks>
		Private Function CreateExcelXmlDocument() As XmlTextWriter

			Dim l_xw As XmlTextWriter = GetXmlWriter

			l_xw.Namespaces = True
			l_xw.Formatting = Formatting.Indented
			l_xw.Indentation = 1

			l_xw.WriteStartDocument()
			l_xw.WriteProcessingInstruction("mso-application", "progid='Excel.Sheet'")

			l_xw.WriteStartElement("Workbook", SPREADSHEET_NAMESPACE)

			l_xw.WriteStartElement("DocumentProperties", OFFICE_NAMESPACE)

			l_xw.WriteStartElement("Author", OFFICE_NAMESPACE)
			l_xw.WriteString(Me.GetType.Assembly.FullName)
			l_xw.WriteEndElement()

			l_xw.WriteStartElement("LastAuthor")

			If Not Environment.UserDomainName = "" Then
				l_xw.WriteString(Environment.UserDomainName & "\" & Environment.UserName)
			Else
				l_xw.WriteString(Environment.UserName)
			End If

			l_xw.WriteEndElement()

			l_xw.WriteStartElement("Created")
			l_xw.WriteString(DateTime.Now.ToString("s") & "Z")
			l_xw.WriteEndElement()

			l_xw.WriteStartElement("Company")
			Dim companyAttr As System.Reflection.AssemblyCompanyAttribute = System.Reflection.Assembly.GetExecutingAssembly.GetCustomAttributes(GetType(System.Reflection.AssemblyCompanyAttribute), False)(0)
			l_xw.WriteString(companyAttr.Company)
			l_xw.WriteEndElement()

			l_xw.WriteStartElement("Version")
			l_xw.WriteString("11")
			l_xw.WriteEndElement()

			l_xw.WriteEndElement()

			l_xw.WriteStartElement("OfficeDocumentSettings", OFFICE_NAMESPACE)
			l_xw.WriteElementString("DownloadComponents", "")
			l_xw.WriteElementString("LocationOfComponents", "")
			l_xw.WriteEndElement()

			Return l_xw

		End Function

		''' <summary>
		''' Private Method to Add Excel Workbook Options to the Xml Writer.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <remarks></remarks>
		Private Sub AddExcelWorkbookOptions( _
			ByRef xw As XmlWriter _
		)

			xw.WriteStartElement("WorksheetOptions", EXCEL_NAMESPACE)

			xw.WriteStartElement("Print")
			xw.WriteElementString("ValidPrinterInfo", "")
			xw.WriteElementString("PaperSizeIndex", "9")
			xw.WriteElementString("HorizontalResolution", "600")
			xw.WriteElementString("VerticalResolution", "600")
			xw.WriteEndElement()

			xw.WriteElementString("Selected", "")

			xw.WriteStartElement("Panes")

			xw.WriteStartElement("Pane")
			xw.WriteElementString("Number", "1")
			xw.WriteElementString("ActiveRow", "1")
			xw.WriteElementString("ActiveCol", "1")
			xw.WriteEndElement()

			xw.WriteEndElement()

			xw.WriteElementString("ProtectObjects", "False")

			xw.WriteElementString("ProtectScenarios", "False")

			xw.WriteEndElement()

			xw.WriteEndElement()

		End Sub

		''' <summary>
		''' Private Method to Add a Row to the Excel Worksheet.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <param name="list"></param>
		''' <param name="styleName">The Style Name to Use in Writing the Row.</param>
		''' <remarks></remarks>
		Private Sub AddExcelRow( _
			ByRef xw As XmlWriter, _
			ByVal list As IList, _
			ByVal styleName As String _
		)

			xw.WriteStartElement("Row")
			For i As Integer = 0 To list.Count - 1
				WriteExcelDataCell(xw, list(i), styleName)
			Next
			xw.WriteEndElement()

		End Sub

		''' <summary>
		''' Private Function to Compare the Equality of Two Property Types.
		''' </summary>
		''' <param name="prop1">The First Property to Compare.</param>
		''' <param name="prop2">The Second Property to Compare.</param>
		''' <returns>A Boolean Value indicating whether the Properties are Functionally Equal.</returns>
		''' <remarks></remarks>
		Private Function CompareProperties( _
			ByVal prop1 As PropertyInfo, _
			ByVal prop2 As PropertyInfo _
		) As Boolean

			Return prop1.Name = prop2.Name AndAlso prop1.ReflectedType Is prop2.ReflectedType AndAlso prop1.DeclaringType Is prop2.DeclaringType AndAlso prop1.PropertyType Is prop2.PropertyType
		
		End Function

		''' <summary>
		''' Private Function to Write a Header Cell to the Excel Worksheet.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <param name="prop">The Property to Write the Header Cell For.</param>
		''' <param name="stylename">The Style Name to Use in Writing the Cell.</param>
		''' <param name="prefix">The Property Prefix to Use.</param>
		''' <remarks></remarks>
		Private Sub WriteExcelHeaderCell( _
			ByRef xw As XmlWriter, _
			ByVal prop As PropertyInfo, _
			ByVal stylename As String, _
			ByVal prefix As String _
		)

			xw.WriteStartElement("Cell")
			xw.WriteAttributeString("ss", "StyleID", SPREADSHEET_NAMESPACE, stylename)
			xw.WriteStartElement("Data")
			xw.WriteAttributeString("ss", "Type", SPREADSHEET_NAMESPACE, "String")
			If prefix = "" Then
				xw.WriteString(Convert.ToString(CamelCaseWords(prop.Name)))
			Else
				xw.WriteString(Convert.ToString(prefix & "." & CamelCaseWords(prop.Name)))
			End If
			xw.WriteEndElement()
			xw.WriteEndElement()

		End Sub

		''' <summary>
		''' Private Function to Get an Excel Data Type for a System Data Type.
		''' </summary>
		''' <param name="type">The System Data Type to Convert.</param>
		''' <returns>A String Representing an Excel Data Type.</returns>
		''' <remarks></remarks>
		Private Function GetExcelDataTypeName( _
			ByVal type As System.Type _
		) As String
		
			If type Is GetType(String) Then
				Return "String"
			ElseIf type Is GetType(Int16) Then
				Return "Number"
			ElseIf type Is GetType(Int32) Then
				Return "Number"
			ElseIf type Is GetType(Int64) Then
				Return "Number"
			ElseIf type Is GetType(Integer) Then
				Return "Number"
			ElseIf type Is GetType(DateTime) Then
				Return "String"
			Else
				Return "String"
			End If

		End Function

		''' <summary>
		''' Private Function to Write a Excel Cell to the Excel Worksheet.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <param name="obj">The Object to Write the Cell for.</param>
		''' <param name="styleName">The Style Name to Use in Writing the Cell.</param>
		''' <remarks></remarks>
		Private Sub WriteExcelDataCell( _
			ByRef xw As XmlWriter, _
			ByVal obj As Object, _
			ByVal styleName As String _
		)

			If IsNothing(obj) Then

				WriteExcelDataCell(xw)

			Else

				xw.WriteStartElement("Cell")
				If obj.GetType() Is GetType(Uri) Then
					xw.WriteAttributeString("ss", "HRef", SPREADSHEET_NAMESPACE, _
					CType(obj, Uri).ToString)
					xw.WriteAttributeString("ss", "StyleID", SPREADSHEET_NAMESPACE, "Hyperlink")
				Else
					xw.WriteAttributeString("ss", "StyleID", SPREADSHEET_NAMESPACE, styleName)
				End If

				xw.WriteStartElement("Data")
				xw.WriteAttributeString("ss", "Type", SPREADSHEET_NAMESPACE, _
				GetExcelDataTypeName(obj.GetType))

				Dim analyser As TypeAnalyser = _
				TypeAnalyser.GetInstance(obj.GetType)

				If analyser.IsICollection Then

					Dim coll As ICollection = obj

					If coll.Count > 0 Then

						Dim sb As New System.Text.StringBuilder
						Dim isFirst As Boolean = True

						For Each singleObj As Object In coll

							If Not isFirst Then
							sb.Append(Environment.NewLine)
							End If
							isFirst = False
							sb.Append(singleObj.ToString)
				
						Next

						xw.WriteString(sb.ToString)
				
					End If
				
				ElseIf obj.GetType Is GetType(Uri) Then
					
					xw.WriteString("Link")

				Else
				
					xw.WriteString(Convert.ToString(obj))

				End If

				xw.WriteEndElement()
				xw.WriteEndElement()

			End If

		End Sub

		''' <summary>
		''' Private Function to Write a Excel Data Cell to the Excel Worksheet.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <remarks></remarks>
		Private Sub WriteExcelDataCell( _
			ByRef xw As XmlWriter _
		)
			xw.WriteStartElement("Cell")
			xw.WriteStartElement("Data")
			xw.WriteAttributeString("ss", "Type", SPREADSHEET_NAMESPACE, "String")
			xw.WriteEndElement()
			xw.WriteEndElement()
		End Sub

		''' <summary>
		''' Private Function to Write a Style to the Excel Workbook.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <param name="styleId">The Id of the Style to Write.</param>
		''' <param name="styleName">The Name of the Style to Write.</param>
		''' <param name="vAlign">The Vertical Align of the Style to Write.</param>
		''' <param name="hAlign">The Horizontal Align of the Style to Write.</param>
		''' <param name="fontFamily">The Font Family of the Style to Write.</param>
		''' <param name="fontBold">The Font Bold Style of the Style to Write.</param>
		''' <remarks></remarks>
		Private Sub WriteStyle( _
			ByRef xw As XmlWriter, _
			ByVal styleId As String, _
			ByVal styleName As String, _
			Optional ByVal vAlign As VerticalAlignment = VerticalAlignment.Bottom, _
			Optional ByVal hAlign As HorizontalAlignment = HorizontalAlignment.Left, _
			Optional ByVal fontFamily As String = "Arial", _
			Optional ByVal fontBold As Boolean = False, _
			Optional ByVal fontUnderline As Boolean = False, _
			Optional ByVal fontColour As String = Nothing _
		)

			xw.WriteStartElement("Style")
			xw.WriteAttributeString("ss", "ID", SPREADSHEET_NAMESPACE, styleId)
			xw.WriteAttributeString("ss", "Name", SPREADSHEET_NAMESPACE, styleName)

			xw.WriteStartElement("Alignment")
			xw.WriteAttributeString("ss", "Vertical", SPREADSHEET_NAMESPACE, vAlign.ToString)
			xw.WriteAttributeString("ss", "Horizontal", SPREADSHEET_NAMESPACE, hAlign.ToString)
			xw.WriteEndElement()

			xw.WriteElementString("Borders", "")

			xw.WriteStartElement("Font")

			If Not fontFamily = "Arial" Then _
				xw.WriteAttributeString("x", "Family", EXCEL_NAMESPACE, fontFamily)

			If fontBold Then
				xw.WriteAttributeString("ss", "Bold", SPREADSHEET_NAMESPACE, "1")
			Else
				xw.WriteAttributeString("ss", "Bold", SPREADSHEET_NAMESPACE, "0")
			End If

			If fontUnderline Then _
				xw.WriteAttributeString("ss", "Underline", SPREADSHEET_NAMESPACE, "Single")

			If Not fontColour = Nothing Then _
			xw.WriteAttributeString("ss", "Color", SPREADSHEET_NAMESPACE, fontColour)

			xw.WriteEndElement()

			xw.WriteElementString("Interior", "")
			xw.WriteElementString("NumberFormat", "")
			xw.WriteElementString("Protection", "")
			xw.WriteEndElement()

		End Sub

		''' <summary>
		''' Private Function to Write a Workbook.
		''' </summary>
		''' <param name="xw">The Xml Writer to Write To.</param>
		''' <param name="wHeight">The Window Height to Write.</param>
		''' <param name="wWidth">The Window Width to Write.</param>
		''' <param name="topX">The Window Top X to Write.</param>
		''' <param name="topY">The Window Top Y to Write.</param>
		''' <param name="protectStructure">A Boolean Value indicating whether the Structure should be protected.</param>
		''' <param name="protectWindows">A Boolean Value indicating whether the Windows should be protected.</param>
		''' <remarks></remarks>
		Private Sub WriteWorkbook( _
			ByRef xw As XmlWriter, _
			Optional ByVal wHeight As Integer = 8000, _
			Optional ByVal wWidth As Integer = 16000, _
			Optional ByVal topX As Integer = 0, _
			Optional ByVal topY As Integer = 45, _
			Optional ByVal protectStructure As Boolean = False, _
			Optional ByVal protectWindows As Boolean = False _
		)

			xw.WriteStartElement("ExcelWorkbook", EXCEL_NAMESPACE)
			xw.WriteElementString("WindowHeight", wHeight.ToString)
			xw.WriteElementString("WindowWidth", wWidth.ToString)
			xw.WriteElementString("WindowTopX", topX.ToString)
			xw.WriteElementString("WindowTopY", topY.ToString)
			xw.WriteElementString("ProtectStructure", CamelCaseWords(protectStructure.ToString))
			xw.WriteElementString("ProtectWindows", CamelCaseWords(protectWindows.ToString))
			xw.WriteEndElement()

		End Sub

	#End Region

	#Region " Public Methods "

		''' <summary>
		''' Public Method to Save an IList of Data to an Excel Worksheet.
		''' </summary>
		''' <param name="propertylist">The Properties that should be in the Header Row.</param>
		''' <param name="dataList">The IList of Data Rows.</param>
		''' <param name="listName">The Name of the List.</param>
		''' <returns>A Boolean to indicate whether the List has been successfully written.</returns>
		''' <remarks></remarks>
		Public Function SaveListToExcel( _
			ByVal propertylist As IList, _
			ByVal dataList As IList, _
			Optional ByVal listName As String = "List" _
		) As Boolean

			Try

				Dim l_xw As XmlWriter = CreateExcelXmlDocument()

				WriteWorkbook(l_xw)

				l_xw.WriteStartElement("Styles")

				WriteStyle(l_xw, "Default", "Normal")
				WriteStyle(l_xw, "Header", "HeaderStyle", VerticalAlignment.Bottom, HorizontalAlignment.Center, _
					"Swiss", True)
				WriteStyle(l_xw, "Hyperlink", "Hyperlink", VerticalAlignment.Bottom, HorizontalAlignment.Left, _
					"Swiss", False, True, "#0000FF")

				l_xw.WriteEndElement()

				l_xw.WriteStartElement("Worksheet")
				l_xw.WriteAttributeString("ss", "Name", SPREADSHEET_NAMESPACE, listName)

				l_xw.WriteStartElement("Table")
				l_xw.WriteAttributeString("ss", "ExpandedRowCount", SPREADSHEET_NAMESPACE, dataList.Count + 1)
				l_xw.WriteAttributeString("x", "FullColumns", SPREADSHEET_NAMESPACE, "1")
				l_xw.WriteAttributeString("x", "FullRows", SPREADSHEET_NAMESPACE, "1")

				AddExcelRow(l_xw, propertylist, "Header")

				For i As Integer = 0 To dataList.Count - 1
					AddExcelRow(l_xw, dataList(i), "Default")
				Next

				l_xw.WriteEndElement()

				l_xw.WriteEndElement()

				l_xw.WriteEndDocument()

				l_xw.Flush()
				l_xw.Close()

				Return True

			Catch ex As Exception
			
				Return False
			
			End Try

		End Function

	#End Region

	End Class

End Namespace