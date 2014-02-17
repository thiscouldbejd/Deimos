Imports Deimos.Visio.ShapeFormatter
Imports Leviathan.Comparison.Comparer

Namespace Visio

	Public Class ShapeManipulation

		#Region " Private Shared Functions "

			''' <summary>
			''' Private Method to get a cell from a shape.
			''' </summary>
			''' <param name="sh">The Shape to Get the Cell from.</param>
			''' <param name="cellName">The Name of the Cell (without prefix, case-insensitive).</param>
			''' <returns>A Cell or Nothing.</returns>
			''' <remarks></remarks>
			Private Shared Function GetCell( _
				ByVal sh As V.Shape, _
				ByVal cellName As String, _
				ByRef cellType As System.Type _
			) As V.Cell

				If sh.SectionExists(V.VisSectionIndices.visSectionProp, True) OrElse _
					sh.SectionExists(V.VisSectionIndices.visSectionProp, False) Then

					Dim propertySection As V.Section = _
						sh.Section(V.VisSectionIndices.visSectionProp)

					If Not propertySection Is Nothing Then

						Dim propertySectionCount As Integer = propertySection.Count - 1

						For i As Integer = 0 To propertySectionCount

							Dim bolFoundCell As Boolean

							If String.Compare(cellName, propertySection(i).Name, True) = 0 Then

								bolFoundCell = True

							Else

								If propertySection(i).Count >= V.VisCellIndices.visCustPropsLabel Then

									If String.Compare(cellName, propertySection(i).Cell(V.VisCellIndices.visCustPropsLabel).Formula, True) = 0 Then

										bolFoundCell = True

									End If

								End If

							End If

							If bolFoundCell Then

								If propertySection(i).Count >= V.VisCellIndices.visCustPropsValue Then

									If propertySection(i).Count >= V.VisCellIndices.visCustPropsType Then

										Dim strPropType As String = propertySection(i).Cell(V.VisCellIndices.visCustPropsType).Formula.Trim(QUOTE_DOUBLE).Trim(SPACE)

										If strPropType = "2" Then

											cellType = GetType(Single)

										ElseIf strPropType = "3" Then

											cellType = GetType(Boolean)

										ElseIf strPropType = "5" Then

											cellType = GetType(DateTime)

										Else

											cellType = GetType(String)

										End If

									End If

									Return propertySection(i).Cell(V.VisCellIndices.visCustPropsValue)

								Else

									Exit For

								End If

							End If

						Next

					End If

				End If

				Return Nothing

			End Function

			''' <summary>
			''' Public Method to Get an Array of Shapes from a Shapes Class.
			''' </summary>
			''' <param name="shapes"></param>
			''' <returns></returns>
			''' <remarks></remarks>
			Private Shared Function GetShapes( _
				ByVal shapes As V.Shapes _
			) As V.Shape()

				Dim shapes_Count As Integer = shapes.Count

				Dim returnArray As V.Shape() = Array.CreateInstance(GetType(V.Shape), shapes_Count)

				For i As Integer = 1 To shapes_Count

					returnArray(i - 1) = shapes(i)

				Next

				Return returnArray

			End Function

		#End Region

		#Region " Public Shared Functions "

			''' <summary>
			''' Public Method to Get an Array of Shapes from a Parent Page.
			''' </summary>
			''' <param name="pg"></param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Shared Function GetShapes( _
				ByVal pg As V.Page _
			) As V.Shape()

				Return GetShapes(pg.Shapes)

			End Function

			''' <summary>
			''' Public Method to Get an Array of Shapes from a Parent Shape.
			''' </summary>
			''' <param name="sh"></param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Shared Function GetShapes( _
				ByVal sh As V.Shape _
			) As V.Shape()

				Return GetShapes(sh.Shapes)

			End Function

			''' <summary>
			''' Public Method to Get a Shapes Rows.
			''' </summary>
			''' <param name="sh">The Shape to Get the Rows for.</param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Shared Function GetRows( _
				ByVal sh As V.Shape _
			) As V.Row()

				Dim retList As New ArrayList

				If (0 <> sh.SectionExists(V.VisSectionIndices.visSectionProp, _
					CShort(V.VisExistsFlags.visExistsAnywhere))) Then

					Dim propertySection As V.Section = _
						sh.Section(V.VisSectionIndices.visSectionProp)

					If Not propertySection Is Nothing Then

						Dim propertySectionCount As Integer = propertySection.Count - 1

						For i As Integer = 0 To propertySectionCount

							retList.Add(propertySection(i))

						Next

					End If

				End If

				Return retList.ToArray(GetType(V.Row))

			End Function

			''' <summary>
			''' Public Method to Get a Row Formula.
			''' </summary>
			''' <param name="rw">The Row to Get the Formula for.</param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Shared Function GetRowFormula( _
				ByVal rw As V.Row _
			) As String

				If rw.Count >= V.VisCellIndices.visUserValue Then _
					Return rw.Cell(V.VisCellIndices.visUserValue).Formula.Trim(QUOTE_DOUBLE).Trim(SPACE)

				Return Nothing

			End Function

			''' <summary>
			''' Public Method to Get a Cell Formula.
			''' </summary>
			''' <param name="sh">The Shape to Get the Cell Formula for.</param>
			''' <param name="cellName">The Name of the Cell.</param>
			''' <param name="cellFormula">The Formula of the Cell as an Out Parameter.</param>
			''' <returns>A Boolean if the Get is Successful.</returns>
			''' <remarks></remarks>
			Public Shared Function GetCellFormula( _
				ByVal sh As V.Shape, _
				ByVal cellName As String, _
				ByRef cellFormula As Object _
			) As Boolean

				Dim shape_Cell_Type As Type = Nothing
				Dim shape_Cell As V.Cell = GetCell(sh, cellName, shape_Cell_Type)

				If Not shape_Cell Is Nothing Then

					cellFormula = New FromString().Parse(shape_Cell.Formula.Trim(QUOTE_DOUBLE).Trim(SPACE), New Boolean, shape_Cell_Type)

					Return True

				Else

					Return False

				End If

			End Function

			''' <summary>
			''' Public Method to Set a Cell Formula.
			''' </summary>
			''' <param name="sh">The Shape to Set the Cell Formula for.</param>
			''' <param name="cellName">The Name of the Cell.</param>
			''' <param name="cellFormula">The Formula of the Cell.</param>
			''' <returns>A Boolean if the Set is Successful.</returns>
			''' <remarks></remarks>
			Public Shared Function SetCellFormula( _
				ByVal sh As V.Shape, _
				ByVal cellName As String, _
				ByVal cellFormula As Object _
			) As Boolean

				If Not IsNothing(cellFormula) Then

					Dim shape_Cell_Type As Type = Nothing
					Dim shape_Cell As V.Cell = GetCell(sh, cellName, shape_Cell_Type)

					If Not shape_Cell Is Nothing Then

						If shape_Cell_Type Is GetType(String) Then

							shape_Cell.Formula = ShapeData.StringToFormulaForString(cellFormula.ToString)

						Else

							shape_Cell.Formula = cellFormula.ToString

						End If

						Return True

					End If

				End If

				Return False

			End Function

		#End Region

	End Class

End Namespace