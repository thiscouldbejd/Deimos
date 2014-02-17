Imports Deimos.Visio.CellConstants
Imports V = Microsoft.Office.Interop.Visio

Namespace Visio

	Public Class ShapeFormatter

		#Region " Private Formatting Methods "

			''' <summary>
			''' Method to set a batch of Shape Properties.
			''' </summary>
			''' <param name="sh">The Shape to Set.</param>
			''' <param name="properties">An Array of Properties.</param>
			''' <param name="values">An Array of Formulas.</param>
			''' <param name="childLevels">The Number of Levels to Drill-Down to.</param>
			''' <remarks></remarks>
			Private Sub SetShapeProperties( _
				ByVal sh As V.Shape, _
				ByVal properties As Short(), _
				ByVal values As Object(), _
				Optional ByVal childLevels As Integer = 0 _
			)

				sh.SetFormulas(properties, values, V.VisGetSetArgs.visSetUniversalSyntax)

				If childLevels > 0 Then

					Dim child_Shapes As V.Shapes = sh.Shapes
					Dim child_ShapeCount As Integer = child_Shapes.Count

					For i As Integer = 1 To child_ShapeCount

						SetShapeProperties(child_Shapes(i), properties, values, (childLevels - 1))

					Next

				End If

			End Sub

			''' <summary>
			''' Sets a Shapes Foreground Colour.
			''' </summary>
			''' <param name="sh">The Shape to Set.</param>
			''' <param name="red">The Red Value.</param>
			''' <param name="green">The Green Value.</param>
			''' <param name="blue">The Blue Value.</param>
			''' <param name="transparency">The Transparency Value.</param>
			''' <remarks></remarks>
			Private Sub SetShapeForeground( _
				ByVal sh As V.Shape, _
				ByVal red As Byte, _
				ByVal green As Byte, _
				ByVal blue As Byte, _
				Optional ByVal transparency As Byte = 255 _
			)

				Dim properties(5) As Short

				properties(0) = CShort(V.VisSectionIndices.visSectionObject)
				properties(1) = CShort(V.VisRowIndices.visRowFill)
				properties(2) = CShort(V.VisCellIndices.visFillForegnd)

				properties(3) = CShort(V.VisSectionIndices.visSectionObject)
				properties(4) = CShort(V.VisRowIndices.visRowFill)
				properties(5) = CShort(V.VisCellIndices.visFillForegndTrans)

				Dim formulas(1) As Object

				formulas(0) = ShapeData.StringToFormulaForString(GenerateRGBValue(red, green, blue))
				formulas(1) = ShapeData.StringToFormulaForString(GenerateTransparencyValue(transparency))

				SetShapeProperties(sh, properties, formulas, Integer.MaxValue)

			End Sub

			''' <summary>
			''' Sets a Shapes Background Colour.
			''' </summary>
			''' <param name="sh">The Shape to Set.</param>
			''' <param name="red">The Red Value.</param>
			''' <param name="green">The Green Value.</param>
			''' <param name="blue">The Blue Value.</param>
			''' <param name="transparency">The Transparency Value.</param>
			''' <remarks></remarks>
			Private Sub SetShapeBackground( _
				ByVal sh As V.Shape, _
				ByVal red As Byte, _
				ByVal green As Byte, _
				ByVal blue As Byte, _
				Optional ByVal transparency As Byte = 255 _
			)

				Dim properties(5) As Short

				properties(0) = CShort(V.VisSectionIndices.visSectionObject)
				properties(1) = CShort(V.VisRowIndices.visRowFill)
				properties(2) = CShort(V.VisCellIndices.visFillBkgnd)

				properties(3) = CShort(V.VisSectionIndices.visSectionObject)
				properties(4) = CShort(V.VisRowIndices.visRowFill)
				properties(5) = CShort(V.VisCellIndices.visFillBkgndTrans)

				Dim formulas(1) As Object

				formulas(0) = ShapeData.StringToFormulaForString(GenerateRGBValue(red, green, blue))
				formulas(1) = ShapeData.StringToFormulaForString(GenerateTransparencyValue(transparency))

				SetShapeProperties(sh, properties, formulas, Integer.MaxValue)

			End Sub

		#End Region

		#Region " Public Formatting Methods "

			Public Sub MakeShapeGreen( _
				ByVal sh As V.Shape _
			)

				SetShapeForeground(sh, 0, 255, 0, 75)

			End Sub

			Public Sub MakeShapeRed( _
				ByVal sh As V.Shape _
			)

				SetShapeForeground(sh, 255, 0, 0, 75)

			End Sub

			Public Sub TransformShapeText( _
				ByVal shape As V.Shape, _
				ByVal widthScalar As Single, _
				ByVal heightScalar As Single, _
				ByVal angle As Single, _
				ByVal pinXScalar As Single, _
				ByVal pinYScalar As Single, _
				ByVal locPinXScalar As Single, _
				ByVal locPinYScalar As Single _
			)

				' Width Scalar
				shape.Cells(NAME_TEXT_WIDTH).Formula = _
				NAME_WIDTH & ASTERISK & widthScalar.ToString

				' Height Scalar
				shape.Cells(NAME_TEXT_HEIGHT).Formula = _
				NAME_HEIGHT & ASTERISK & heightScalar.ToString

				' Angle
				shape.Cells(NAME_TEXT_ANGLE).Formula = _
				angle.ToString & " deg"

				' PinX
				shape.Cells(NAME_TEXT_PIN_X).Formula = _
				NAME_WIDTH & ASTERISK & pinXScalar.ToString

				' PinY
				shape.Cells(NAME_TEXT_PIN_Y).Formula = _
				NAME_HEIGHT & ASTERISK & pinYScalar.ToString

				' LocPinX
				shape.Cells(NAME_TEXT_LOC_PIN_X).Formula = _
				NAME_TEXT_WIDTH & ASTERISK & locPinXScalar.ToString

				' LocPinY
				shape.Cells(NAME_TEXT_LOC_PIN_Y).Formula = _
				NAME_TEXT_HEIGHT & ASTERISK & locPinYScalar.ToString

			End Sub

		#End Region

		#Region " Public Shared Methods "

			Public Shared Function GenerateRGBValue( _
				ByVal red As Byte, _
				ByVal green As Byte, _
				ByVal blue As Byte _
			) As String

				Dim sb As New System.Text.StringBuilder()

				sb.Append(ColourConstants.PREFIX_RGB)
				sb.Append(BRACKET_START)
				sb.Append(red.ToString)
				sb.Append(COMMA)
				sb.Append(green.ToString)
				sb.Append(COMMA)
				sb.Append(blue.ToString)
				sb.Append(BRACKET_END)

				Return sb.ToString

			End Function

			Public Shared Function GenerateTransparencyValue( _
				ByVal transparency As Byte _
			) As String

				Return transparency.ToString & PERCENTAGE_MARK

			End Function

		#End Region

	End Class

End Namespace