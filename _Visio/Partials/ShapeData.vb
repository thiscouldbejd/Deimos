Imports V = Microsoft.Office.Interop.Visio

Namespace Visio

	Public Class ShapeData

		Private Shared ShapeCache As New Hashtable

		#Region " Public Shared Functions "

			Public Shared Function StringToFormulaForString( _
				ByVal formula As String _
			) As String

				Return QUOTE_DOUBLE & formula.Replace(QUOTE_DOUBLE, _
					QUOTE_DOUBLE & QUOTE_DOUBLE) & QUOTE_DOUBLE

			End Function

			Public Shared Function GetShapesWithProperties( _
				ByVal page As V.Page, _
				ByRef returnedShapeIds As Short(), _
				ByRef returnedShapeProperties As String()(), _
				Optional ByVal shapeTypeName As String = Nothing _
			) As Boolean

				Return GetShapesWithProperties(page.Shapes, _
					returnedShapeIds, returnedShapeProperties, shapeTypeName)

			End Function

			Public Shared Function GetShapesWithProperties( _
				ByVal shapes As V.Shapes, _
				ByRef returnedShapeIds As Short(), _
				ByRef returnedShapeProperties As String()(), _
				Optional ByVal shapeTypeName As String = Nothing _
			) As Boolean

				Dim shapes_Count As Integer = shapes.Count

				returnedShapeIds = Array.CreateInstance(GetType(Short), shapes_Count)
				returnedShapeProperties = Array.CreateInstance(GetType(String).MakeArrayType, shapes_Count)

				Dim sectionIndex As Short = CShort(V.VisSectionIndices.visSectionProp)

				Dim index As Integer = 0

				For i As Integer = 1 To shapes_Count

					Dim page_Shape As V.Shape = shapes(i)

					If (0 <> page_Shape.SectionExists(sectionIndex, _
						CShort(V.VisExistsFlags.visExistsAnywhere))) Then

						Dim shapeName As String = page_Shape.Name

						If shapeName.Contains(FULL_STOP) Then _
							shapeName = shapeName.Substring(0, shapeName.IndexOf(FULL_STOP))

						If String.IsNullOrEmpty(shapeTypeName) OrElse _
							String.Compare(shapeName, shapeTypeName, True) = 0 Then

							Dim shapeProperties As String()

							If Not String.IsNullOrEmpty(shapeTypeName) _
								AndAlso ShapeCache.Contains(shapeTypeName) Then

								shapeProperties = ShapeCache(shapeTypeName)

							Else

								Dim propertySection As V.Section = _
									page_Shape.Section(V.VisSectionIndices.visSectionProp)

								shapeProperties = Array.CreateInstance(GetType(String), propertySection.Count)

								For j As Integer = 0 To propertySection.Count - 1

									shapeProperties(j) = propertySection(j).Name

								Next

								If Not String.IsNullOrEmpty(shapeTypeName) Then _
									ShapeCache.Add(shapeTypeName, shapeProperties)

							End If

							returnedShapeIds(index) = CShort(page_Shape.ID)
							returnedShapeProperties(index) = shapeProperties
							index += 1

						End If

					End If

				Next

				Array.Resize(returnedShapeIds, index)
				Array.Resize(returnedShapeProperties, index)

				Return returnedShapeIds.Length > 0

			End Function

		#End Region

	End Class

End Namespace