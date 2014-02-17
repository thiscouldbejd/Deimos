Imports Leviathan.Visualisation

Namespace Commands

	Partial Public Class WordCommands
	
		#Region " Public Command Methods "
		
			<Command( _
				ResourceContainingType:=GetType(WordCommands), _
				ResourceName:="CommandDetails", _
				Name:="output", _
				Description:="@commandWordDescriptionOutput@" _
			)> _
			Public Sub ProcessCommandOutput( _
				<Configurable( _
					ResourceContainingType:=GetType(WordCommands), _
					ResourceName:="CommandDetails", _
					Description:="@commandWordParameterDescriptionFormattedObjects@" _
				)> _
				ByVal value As Cube _
			)
			
				If Not value Is Nothing Then Document.OutputCube(New Cube() {value}, host)
				
			End Sub
			
		#End Region
	
	End Class
	
End Namespace