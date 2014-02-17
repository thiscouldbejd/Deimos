Imports V = Microsoft.Office.Interop.Visio

Namespace Visio

	Public Class ShapeInteraction

		#Region " Public Shared Functions "

			Public Shared Function GetClickedShape( _
				ByVal visioPage As V.Page, _
				ByVal x As Double, _
				ByVal y As Double _
			) As V.Shape

				Return (GetClickedShape(visioPage, x, y, 0.0001))

			End Function

			Public Shared Function GetClickedShape( _
				ByVal visioPage As V.Page, _
				ByVal x As Double, _
				ByVal y As Double, _
				ByVal tolerance As Double _
			) As V.Shape

				Try
					Dim visSelection As V.Selection = visioPage.SpatialSearch(x, y, _
						CType(V.VisSpatialRelationCodes.visSpatialContainedIn, Short), _
						tolerance, V.VisSpatialRelationFlags.visSpatialFrontToBack)

					If visSelection.Count > 0 Then
						Return CType(visSelection(1), V.Shape)
					Else
						Return Nothing
					End If

				Catch ex As Exception
					Return Nothing
				End Try

			End Function

		#End Region

	End Class

End Namespace