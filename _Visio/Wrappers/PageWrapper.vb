Imports Deimos.Visio.CellConstants
Imports Leviathan
Imports Leviathan.Caching
Imports Leviathan.Commands
Imports Leviathan.Inspection.TypeAnalyser
Imports Leviathan.Visualisation
Imports System.Data
Imports System.Runtime.InteropServices

Namespace Visio

	Partial Public Class PageWrapper
	
		#Region " Private Constants "
		
			''' <summary>
			''' Public Constant Reference to the Name of the SnapshotShapeCellValue Method.
			''' </summary>
			''' <remarks></remarks>
			Private Const METHOD_SNAPSHOTSHAPECELLVALUE As String = "SnapshotShapeCellValue"
			
			''' <summary>
			''' Public Constant Reference to the Name of the RestoreShapeCellValue Method.
			''' </summary>
			''' <remarks></remarks>
			Private Const METHOD_RESTORESHAPECELLVALUE As String = "RestoreShapeCellValue"
			
		#End Region
		
		#Region " Private Methods "
		
			Private Sub SnapshotShapeCellValue( _
				ByRef sh As V.Shape, _
				ByVal ParamArray cellNames As String() _
			)
			
				For Each cellName As String In cellNames
				
					SnapshotShapeCellValue(sh, cellName)
					
				Next
				
			End Sub
			
			Private Sub SnapshotShapeCellValue( _
				ByRef sh As V.Shape, _
				ByVal cellName As String _
			)
			
				If sh.CellExists(cellName, 0) = -1 Then
				
					Dim cache As Simple = Simple.GetInstance(GetType(PageWrapper).GetHashCode)
					
					cache.Set(sh.Cells(cellName).Formula, METHOD_SNAPSHOTSHAPECELLVALUE.GetHashCode, Page.ID.GetHashCode, sh.ID.GetHashCode, cellName.GetHashCode)
					
				End If
				
				For Each child_sh As V.Shape In sh.Shapes
				
					SnapshotShapeCellValue(child_sh, cellName)
					
				Next
				
			End Sub
			
			Private Sub RestoreShapeCellValue( _
				ByRef sh As V.Shape, _
				ByVal ParamArray cellNames As String() _
			)
			
				For Each cellName As String In cellNames
				
					RestoreShapeCellValue(sh, cellName)
					
				Next
				
			End Sub
			
			Private Sub RestoreShapeCellValue( _
				ByRef sh As V.Shape, _
				ByVal cellName As String _
			)
			
				If sh.CellExists(cellName, 0) = -1 Then
				
					Dim cache As Simple = Simple.GetInstance(GetType(PageWrapper).GetHashCode)
					
					Dim form_Value As String = Nothing
					
					If cache.TryGet(form_Value, METHOD_SNAPSHOTSHAPECELLVALUE.GetHashCode, Page.ID.GetHashCode, sh.ID.GetHashCode, cellName.GetHashCode) Then _
						sh.Cells(cellName).Formula = form_Value
						
				End If
				
				For Each child_sh As V.Shape In sh.Shapes
				
					RestoreShapeCellValue(child_sh, cellName)
					
				Next
				
			End Sub
			
			''' <summary>
			''' Private Method to Output to a Shape/s.
			''' </summary>
			''' <param name="value">The Value to Output.</param>
			''' <param name="shapes">The Shapes to Output To.</param>
			''' <remarks></remarks>
			Private Sub OutputToShape( _
				ByVal value As Object, _
				ByVal shapes As V.Shapes, _
				ByVal page As V.Page, _
				ByRef shape_Ids As ArrayList, _
				ByVal shape_Rows As ArrayList, _
				ByVal shape_Values As ArrayList, _
				Optional ByVal shape_TypeName As String = Nothing _
			)
			
				If IsNothing(value) Then Exit Sub
				
				Dim members As MemberAnalyser() = TypeAnalyser.GetInstance(value.GetType).ExecuteQuery( _
					AnalyserQuery.QUERY_MEMBERS_READABLE.SetPresentAttribute(GetType(UniqueAttribute)))
					
				If members Is Nothing OrElse members.Length = 0 Then Exit Sub
				
				For j As Integer = 0 To members.Length - 1
				
					Dim idMemberName As String = members(j).Name
					Dim idMemberValue As Object = members(j).Read(value)
					
					Dim selectedShape_Ids As Short() = Nothing
					Dim selectedShape_PropertyNames As String()() = Nothing
					
					If Not IsNothing(idMemberValue) AndAlso ShapeData.GetShapesWithProperties(shapes, selectedShape_Ids, selectedShape_PropertyNames, _
						shape_TypeName) Then
						
						Dim idMemberValueString As String = idMemberValue.ToString
						
						Dim formulas As Object() = Array.CreateInstance(GetType(Object), selectedShape_Ids.Length)
						
						Dim SID_SRCStream As Short() = Array.CreateInstance(GetType(Short), formulas.Length * 4)
						
						Dim index As Integer = 0
						
						For i As Integer = 0 To selectedShape_Ids.Length - 1
						
							SID_SRCStream(index) = selectedShape_Ids(i)
							index += 1
							
							SID_SRCStream(index) = V.VisSectionIndices.visSectionProp
							index += 1
							
							SID_SRCStream(index) = CShort(Array.IndexOf(selectedShape_PropertyNames(i), idMemberName))
							index += 1
							
							SID_SRCStream(index) = V.VisCellIndices.visUserValue
							index += 1
							
						Next
						
						page.GetFormulasU(SID_SRCStream, formulas)
						
						For k As Integer = 0 To selectedShape_Ids.Length - 1
						
							If String.Compare(idMemberValueString, formulas(k).ToString.Trim(QUOTE_DOUBLE).Trim(), True) = 0 Then
							
								Dim readableMembers As MemberAnalyser() = TypeAnalyser.GetInstance(value.GetType).ExecuteQuery(AnalyserQuery.QUERY_MEMBERS_READABLE)
								
								For i As Integer = 0 To readableMembers.Length - 1
								
									Dim cellValue As Object = readableMembers(i).Read(value)
									
									If Not IsNothing(cellValue) Then
									
										Dim newCellValueAnalyser As TypeAnalyser = TypeAnalyser.GetInstance(cellValue.GetType)
										
										If newCellValueAnalyser.IsICollection Then
										
											For Each singleCellValue As Object In CType(cellValue, ICollection)
											
												OutputToShape(singleCellValue, shapes.ItemFromID(selectedShape_Ids(k)).Shapes, page, shape_Ids, shape_Rows, shape_Values)
												
											Next
											
										Else
										
											Dim propertyMemberIndex As Integer = Array.IndexOf(selectedShape_PropertyNames(k), readableMembers(i).Name)
											
											If propertyMemberIndex >= 0 Then
											
												shape_Ids.Add(selectedShape_Ids(k))
												shape_Rows.Add(propertyMemberIndex)
												shape_Values.Add(cellValue)
												
											End If
											
										End If
										
									End If
									
								Next
								
							End If
							
						Next
						
					End If
					
				Next
				
			End Sub
			
			''' <summary>
			''' Private Method to Reset Shape/s Cell Formulats
			''' </summary>
			''' <param name="shapes">The Shapes to Output To.</param>
			''' <remarks></remarks>
			Private Sub ClearShapeData( _
				ByVal shapes As V.Shapes, _
				ByVal propertiesToPreserve As String(), _
				ByRef shape_Ids As ArrayList, _
				ByVal shape_Rows As ArrayList, _
				Optional ByVal shape_TypeName As String = Nothing _
			)
			
				Dim selectedShape_Ids As Short() = Nothing
				Dim selectedShape_PropertyNames As String()() = Nothing
				
				If ShapeData.GetShapesWithProperties(shapes, selectedShape_Ids, selectedShape_PropertyNames, shape_TypeName) Then
				
					For i As Integer = 0 To selectedShape_Ids.Length - 1
					
						For j As Integer = 0 To selectedShape_PropertyNames(i).Length - 1
						
							Dim clearProperty As Boolean = True
							
							For k As Integer = 0 To propertiesToPreserve.Length - 1
							
								If String.Compare(selectedShape_PropertyNames(i)(j), propertiesToPreserve(k), True) = 0 Then
								
									clearProperty = False
									Exit For
									
								End If
								
							Next
							
							If clearProperty Then
							
								shape_Ids.Add(selectedShape_Ids(i))
								shape_Rows.Add(j)
								
							End If
							
						Next
						
						ClearShapeData(shapes.ItemFromID(selectedShape_Ids(i)).Shapes, propertiesToPreserve, shape_Ids, shape_Rows)
						
					Next
					
				End If
				
			End Sub
			
		#End Region
		
		#Region " Protected Methods "
		
			Protected Overrides Sub SetPage( _
				ByVal page As Object _
			)
			
				m_Page = page
				
			End Sub
			
		#End Region
		
		#Region " Public Methods "
		
			Public Function GetData( _
				Optional ByVal shapeType As String = Nothing, _
				Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
			) As ICollection
			
				' Create a New DataTable
				Dim dt As New DataTable()
				
				Dim shape_Ids As Short() = Nothing
				Dim property_Names As String()() = Nothing
				
				If ShapeData.GetShapesWithProperties(Page, shape_Ids, property_Names, shapeType) Then
				
					For i As Integer = 0 To property_Names(0).Length - 1
					
						dt.Columns.Add(New System.Data.DataColumn(property_Names(0)(i), GetType(String)))
						
					Next
					
					Dim formulas As Object() = Array.CreateInstance(GetType(Object), shape_Ids.Length * property_Names(0).Length)
					
					Dim SID_SRCStream As Short() = Array.CreateInstance(GetType(Short), formulas.Length * 4)
					
					Dim index As Integer = 0
					
					Dim shapeIdsCount As Integer = shape_Ids.Length
					
					For i As Integer = 0 To shapeIdsCount - 1
					
						For j As Integer = 0 To property_Names(i).Length - 1
						
							SID_SRCStream(index) = shape_Ids(i)
							SID_SRCStream(index + 1) = V.VisSectionIndices.visSectionProp
							SID_SRCStream(index + 2) = CShort(j)
							SID_SRCStream(index + 3) = V.VisCellIndices.visUserValue
							
							index += 4
							
						Next
						
					Next
					
					Page.GetFormulasU(SID_SRCStream, formulas)
					
					For i As Integer = 0 To shapeIdsCount - 1
					
						Dim dr As DataRow = dt.NewRow
						
						For j As Integer = 0 To property_Names(i).Length - 1
						
							dr(j) = formulas((i * property_Names(i).Length) + j).Trim(QUOTE_DOUBLE).Trim(SPACE)
							
						Next
						
						dt.Rows.Add(dr)
						
						If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive) Then Host.Progress(i + 1 / shapeIdsCount, "Interrogating Shapes")
						
					Next
					
				End If
				
				Return GetObjectsFromDataTable(dt)
				
			End Function
			
			Public Function OutputShapes( _
				ByVal shapeType As String, _
				ByVal value As Object(), _
				Optional ByVal host As ICommandsExecution = Nothing _
			) As Boolean
			
				Dim shape_Ids As New ArrayList
				Dim shape_Rows As New ArrayList
				Dim shape_Values As New ArrayList
				
				For i As Integer = 0 To value.Length - 1
				
					OutputToShape(value(i), Page.Shapes, Page, shape_Ids, shape_Rows, shape_Values, shapeType)
					
					If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive) Then Host.Progress(i + 1 / value.Length, "Outputting To Shapes")
					
				Next
				
				' Perform the Set Back
				If shape_Ids.Count > 0 AndAlso shape_Ids.Count = shape_Rows.Count AndAlso shape_Ids.Count = shape_Values.Count Then
				
					Dim formulas As Object() = Array.CreateInstance(GetType(Object), shape_Values.Count)
					
					Dim SID_SRCStream As Short() = Array.CreateInstance(GetType(Short), formulas.Length * 4)
					
					Dim index As Integer = 0
					
					For i As Integer = 0 To shape_Ids.Count - 1
					
						SID_SRCStream(index) = CShort(shape_Ids(i))
						index += 1
						
						SID_SRCStream(index) = V.VisSectionIndices.visSectionProp
						index += 1
						
						SID_SRCStream(index) = CShort(shape_Rows(i))
						index += 1
						
						SID_SRCStream(index) = V.VisCellIndices.visUserValue
						index += 1
						
						If Not shape_Values(i) Is Nothing Then
						
							If TypeAnalyser.IsTypeNumericType(shape_Values(i).GetType) Then
							
								formulas(i) = shape_Values(i).ToString
								
							Else
							
								formulas(i) = ShapeData.StringToFormulaForString(shape_Values(i).ToString)
							
							End If
							
						End If
						
					Next
					
					Return Page.SetFormulas(SID_SRCStream, formulas, V.VisGetSetArgs.visSetBlastGuards) > 0
					
				End If
				
				Return False
				
			End Function
			
			Public Function ResetShapes( _
				ByVal shapeType As String, _
				ByVal shapePropertiesToLeave As String() _
			) As Boolean
			
				Dim shape_Ids As New ArrayList
				Dim shape_Rows As New ArrayList
				
				ClearShapeData(Page.Shapes, shapePropertiesToLeave, shape_Ids, shape_Rows, shapeType)
				
				' Perform the Set Back
				If shape_Ids.Count = shape_Rows.Count Then
				
					Dim formulas As Object() = Array.CreateInstance(GetType(Object), shape_Ids.Count)
					
					Dim SID_SRCStream As Short() = Array.CreateInstance(GetType(Short), formulas.Length * 4)
					
					Dim index As Integer = 0
					
					For i As Integer = 0 To shape_Ids.Count - 1
					
						SID_SRCStream(index) = CShort(shape_Ids(i))
						SID_SRCStream(index + 1) = V.VisSectionIndices.visSectionProp
						SID_SRCStream(index + 2) = CShort(shape_Rows(i))
						SID_SRCStream(index + 3) = V.VisCellIndices.visUserValue
						
						formulas(i) = String.Empty
						
						index += 4
						
					Next
					
					If SID_SRCStream.Length > 0 Then Return Page.SetFormulas(SID_SRCStream, formulas, V.VisGetSetArgs.visSetBlastGuards) > 0
					
				End If
				
				Return False
				
			End Function
			
			Public Sub SnapshotShapeFills()
			
				For Each sh As V.Shape In Page.Shapes
				
					SnapshotShapeCellValue(sh, NAME_FILL_FOREGROUND, NAME_FILL_FOREGROUND_TRANS, NAME_FILL_BACKGROUND, NAME_FILL_BACKGROUND_TRANS, _
						NAME_FILL_PATTERN, NAME_SHADOW_FOREGROUND, NAME_SHADOW_FOREGROUND_TRANS, NAME_SHADOW_BACKGROUND, NAME_SHADOW_BACKGROUND_TRANS, _
						NAME_SHADOW_PATTERN, NAME_SHADOW_OFFSET_X, NAME_SHADOW_OFFSET_Y, NAME_SHADOW_TYPE, NAME_SHADOW_OBLIQUE_ANGLE, NAME_SHADOW_SCALE_FACTOR)
						
				Next
				
			End Sub
			
			Public Sub RestoreShapeFills()
			
				For Each sh As V.Shape In Page.Shapes
				
					RestoreShapeCellValue(sh, NAME_FILL_FOREGROUND, NAME_FILL_FOREGROUND_TRANS, NAME_FILL_BACKGROUND, NAME_FILL_BACKGROUND_TRANS, _
						NAME_FILL_PATTERN, NAME_SHADOW_FOREGROUND, NAME_SHADOW_FOREGROUND_TRANS, NAME_SHADOW_BACKGROUND, NAME_SHADOW_BACKGROUND_TRANS, _
						NAME_SHADOW_PATTERN, NAME_SHADOW_OFFSET_X, NAME_SHADOW_OFFSET_Y, NAME_SHADOW_TYPE, NAME_SHADOW_OBLIQUE_ANGLE, NAME_SHADOW_SCALE_FACTOR)
						
				Next
				
			End Sub
			
			Public Sub OutputCube( _
				ByVal values() As Cube, _
				Optional ByVal host As ICommandsExecution = Nothing _
			)
			
				If Not values Is Nothing Then
				
					For i As Integer = 0 To values.Length - 1
					
						If values(i).HasData AndAlso values(i).LastSlice.HasData Then
						
							If Page.Application.DataFeaturesEnabled Then
							
								Dim dataXml As String = values(i).CreateAdoXml
								
								Dim doc As V.Document = Page.Document
								Dim records As V.DataRecordsets = doc.DataRecordsets
								
								If Not values(i).Title Is Nothing Then
								
									records.AddFromXML(dataXml, 0, Name:=values(i).Title.ToString)
									
								Else
								
									records.AddFromXML(dataXml, 0)
									
								End If
								
							End If
							
						End If
						
					Next
					
				End If
				
			End Sub
			
		#End Region
		
		#Region " Public Shared Methods "
		
			Public Shared Function TryParse( _
				ByVal value As String, _
				ByRef result As PageWrapper _
			) As Boolean
			
				Dim documentName As String = Nothing, pageName As String = Nothing
				
				PageBase.ParseName(value, documentName, pageName)
				
				Dim drawing As DrawingWrapper = Nothing
				
				If DrawingWrapper.TryParse(documentName, drawing) Then
				
					Try
					
						result = New PageWrapper().ParseFromString(pageName, drawing, A.Visio)
						
						result.Parent = drawing
						
						Return True
						
					Catch ex As Exception
					
						If Not drawing Is Nothing Then drawing.Dispose()
						
					Finally
					End Try
					
				End If
				
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
				
					If Not m_Page Is Nothing Then
					
						Marshal.ReleaseComObject(m_Page)
						m_Page = Nothing
						
					End If
					
					If Not Parent Is Nothing Then Parent.Dispose()
					
				End If
				
				disposedValue = True
				
			End Sub
			
		#End Region
		
	End Class
	
End Namespace