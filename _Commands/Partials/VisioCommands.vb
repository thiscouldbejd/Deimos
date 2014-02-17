Imports Leviathan.Visualisation

Namespace Commands

	Partial Public Class VisioCommands
	
		#Region " Public Command Methods "
		
			<Command( _
				ResourceContainingType:=GetType(VisioCommands), _
				ResourceName:="CommandDetails", _
				Name:="input", _
				Description:="@commandVisioDescriptionInput@" _
			)> _
			Public Function ProcessCommandInput( _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionShapeType@" _
				)> _
				ByVal shapeType As String _
			) As ICollection
			
				Return Page.GetData(shapeType, Host)
				
			End Function
			
			<Command( _
				ResourceContainingType:=GetType(VisioCommands), _
				ResourceName:="CommandDetails", _
				Name:="output", _
				Description:="@commandVisioDescriptionOutput@" _
			)> _
			Public Sub ProcessCommandOutput( _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionFormattedObjects@" _
				)> _
				ByVal value As Cube _
			)
			
				If Not value Is Nothing Then Page.OutputCube(New Cube() {value}, Host)
				
			End Sub
			
			<Command( _
				ResourceContainingType:=GetType(VisioCommands), _
				ResourceName:="CommandDetails", _
				Name:="output-shapes", _
				Description:="@commandVisioDescriptionOutputShapes@" _
			)> _
			Public Function ProcessCommandOutputShapes( _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionShapeType@" _
				)> _
				ByVal shapeType As String, _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionObjects@" _
				)> _
				ParamArray ByVal value As Object() _
			) As Boolean
			
				Return Page.OutputShapes(shapeType, value, Host)
				
			End Function
			
			<Command( _
				ResourceContainingType:=GetType(VisioCommands), _
				ResourceName:="CommandDetails", _
				Name:="reset-shapes", _
				Description:="@commandVisioDescriptionResetShapes@" _
			)> _
			Public Function ProcessCommandResetShapes( _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionShapeType@" _
				)> _
				ByVal shapeType As String, _
				<Configurable( _
					ResourceContainingType:=GetType(VisioCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandVisioParameterDescriptionPropertiesToLeave@" _
				)> _
				ParamArray ByVal shapePropertiesToLeave As String() _
			) As Boolean
			
				Return Page.ResetShapes(shapeType, shapePropertiesToLeave)
				
			End Function
			
		#End Region
		
	End Class
	
End Namespace