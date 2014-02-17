Namespace Visio

	Public Class CellConstants

		#Region " General Shared Variables "

			Public Shared PREFIX_USER As String = "User"

			Public Shared PREFIX_PROP As String = "Prop"

			Public Shared NAME_USER_TEMPLATE As String = PREFIX_USER & FULL_STOP & "TemplateName"

			Public Shared NAME_TRANPARENCY As String = "Trans"

			Public Shared NAME_WIDTH As String = "Width"

			Public Shared NAME_HEIGHT As String = "Height"

		#End Region

		#Region " Fill Shared Variables "

			Public Shared NAME_FILL_FOREGROUND As String = "FillForegnd"

			Public Shared NAME_FILL_FOREGROUND_TRANS As String = NAME_FILL_FOREGROUND & NAME_TRANPARENCY

			Public Shared NAME_FILL_BACKGROUND As String = "FillBkgnd"

			Public Shared NAME_FILL_BACKGROUND_TRANS As String = NAME_FILL_BACKGROUND & NAME_TRANPARENCY

			Public Shared NAME_FILL_PATTERN As String = "FillPattern"

		#End Region

		#Region " Shadow Shared Variables "

			Public Shared NAME_SHADOW_FOREGROUND As String = "ShdwForegnd"

			Public Shared NAME_SHADOW_FOREGROUND_TRANS As String = NAME_SHADOW_FOREGROUND & NAME_TRANPARENCY

			Public Shared NAME_SHADOW_BACKGROUND As String = "ShdwBackgnd"

			Public Shared NAME_SHADOW_BACKGROUND_TRANS As String = NAME_SHADOW_BACKGROUND & NAME_TRANPARENCY

			Public Shared NAME_SHADOW_PATTERN As String = "ShdwPattern"

			Public Shared NAME_SHADOW_OFFSET_X As String = "ShapeShdwOffsetX"

			Public Shared NAME_SHADOW_OFFSET_Y As String = "ShapeShdwOffsetY"

			Public Shared NAME_SHADOW_TYPE As String = "ShapeShdwType"

			Public Shared NAME_SHADOW_OBLIQUE_ANGLE As String = "ShapShdwObliqueAngle"

			Public Shared NAME_SHADOW_SCALE_FACTOR As String = "ShapeShdwScaleFactor"

		#End Region

		#Region " Text Shared Variables "

			Public Shared NAME_TEXT_WIDTH As String = "TxtWidth"

			Public Shared NAME_TEXT_HEIGHT As String = "TxtHeight"

			Public Shared NAME_TEXT_ANGLE As String = "TxtAngle"

			Public Shared NAME_TEXT_PIN_X As String = "TxtPinX"

			Public Shared NAME_TEXT_PIN_Y As String = "TxtPinY"

			Public Shared NAME_TEXT_LOC_PIN_X As String = "TxtLocPinX"

			Public Shared NAME_TEXT_LOC_PIN_Y As String = "TxtLocPinY"

		#End Region

		#Region " Public Shared Function "

			Public Shared Function PropertyCellName( _
				ByVal cellName As String _
			) As String

				Return PREFIX_PROP & FULL_STOP & cellName

			End Function

			Public Shared Function UserCellName( _
				ByVal cellName As String _
			) As String

				Return PREFIX_USER & FULL_STOP & cellName

			End Function

		#End Region

	End Class

End Namespace