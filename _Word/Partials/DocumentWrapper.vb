Imports Leviathan.Commands
Imports Leviathan.Visualisation
Imports System.Runtime.InteropServices

Namespace Word

	Partial Public Class DocumentWrapper
	
		#Region " Protected Properties "
		
			Protected ReadOnly Property FullName() As String
				Get
					If Not Document Is Nothing Then Return Document.FullName Else Return Nothing
				End Get
			End Property
			
		#End Region
		
		#Region " Protected Methods "
		
			''' <summary>
			''' Method to Set the Document (overrides base class).
			''' </summary>
			''' <param name="document">The document to set.</param>
			''' <remarks></remarks>
			Protected Overrides Sub SetDocument( _
				ByVal document As Object _
			)
				m_Document = document
			End Sub
			
			Protected Sub WriteSlice( _
				ByVal value As Slice, _
				ByVal columns As List(Of FormatterProperty), _
				Optional ByVal host As Leviathan.Commands.ICommandsExecution = Nothing _
			)
			
				Dim tables As W.Tables = Document.Tables
				
				Dim tb As W.Table = tables.Add(Document.Range, value.Rows.Count + 1, columns.Count, W.WdDefaultTableBehavior.wdWord9TableBehavior, _
					W.WdAutoFitBehavior.wdAutoFitContent)
					
				For j As Integer = 0 To columns.Count - 1
				
					Dim wdCell As W.Cell = tb.Cell(1, j + 1)
					Dim wdRange As W.Range = wdCell.Range
					
					wdRange.Text = ObjectToSingleString(columns(j).DisplayName, " | ")
					
					Marshal.ReleaseComObject(wdRange)
					wdRange = Nothing
					
					Marshal.ReleaseComObject(wdCell)
					wdCell = Nothing
					
				Next
				
				Dim rowCount As Integer = value.Rows.Count
				
				For j As Integer = 0 To rowCount - 1
				
					If Not value.Rows(j) Is Nothing AndAlso value.Rows(j).Cells.Count > 0 Then
					
						For k As Integer = 0 To value.Rows(j).Cells.Count - 1
						
							Dim wdCell As W.Cell = tb.Cell(j + 2, k + 1)
							Dim wdRange As W.Range = wdCell.Range
							
							wdRange.Text = ObjectToSingleString(value.Rows(j)(k), "; ")
							
							Marshal.ReleaseComObject(wdRange)
							wdRange = Nothing
							
							Marshal.ReleaseComObject(wdCell)
							wdCell = Nothing
							
						Next
						
					End If
					
					If Not host Is Nothing AndAlso host.Available(VerbosityLevel.Interactive) Then Host.Progress(j + 1 / rowCount, "Outputting Table Rows")
					
				Next
				
				Marshal.ReleaseComObject(tb)
				tb = Nothing
				
				Marshal.ReleaseComObject(tables)
				tables = Nothing
				
				Save()
				
			End Sub
			
		#End Region
		
		#Region " Public Methods "
		
			Public Sub OutputCube( _
				ByVal values As Cube(), _
				Optional ByVal host As ICommandsExecution = Nothing _
			)
			
				For i As Integer = 0 To values.Length - 1
				
					WriteSlice(values(i).LastSlice, values(i).Columns, host)
					
				Next
				
			End Sub
			
			Public Sub Save()
			
				If Not String.IsNullOrEmpty(Document.Path) Then Document.Save()
				
			End Sub
			
			Public Sub Save( _
				ByVal filePath As String _
			)
			
				If Not String.IsNullOrEmpty(filePath) AndAlso Not IO.File.Exists(filePath) Then Document.SaveAs(filePath)
				
			End Sub
			
			Public Function IsCorrectSpelled( _
				ByVal _value As String _
			) As Boolean
			
				Dim wdApp As W.Application = Document.Application, wdWords As W.Words = Document.Words, wdWord As W.Range = wdWords.First
				
				wdWord.InsertBefore(_value)
				
				Dim wdSErrors As W.ProofreadingErrors = Document.SpellingErrors
				
				Dim retValue As Boolean = (wdSErrors.Count = 0)
				
				If Not wdSErrors Is Nothing Then
				
					Marshal.ReleaseComObject(wdSErrors)
					wdSErrors = Nothing
					
				End If
				
				Marshal.ReleaseComObject(wdWord)
				wdWord = Nothing
				
				Marshal.ReleaseComObject(wdWords)
				wdWords = Nothing
				
				Marshal.ReleaseComObject(wdApp)
				wdApp = Nothing
				
				Dim wdRange As W.Range = Document.Range(0, Document.Characters.Count - 1)
				wdRange.Delete()
				
				Marshal.ReleaseComObject(wdRange)
				wdRange = Nothing
				
				Return retValue
				
			End Function
			
			Public Function GetSpellingErrors( _
				ByVal _value As String, _
				ByVal ParamArray _ignoreWords As String() _
			) As SpellingError()
			
				Dim spellingValue As Integer = 0
				Dim grammarValue As Integer = 0
				
				Dim wdApp As W.Application = Document.Application
				Dim wdWords As W.Words = Document.Words
				Dim wdWord As W.Range = wdWords.First
				
				wdWord.InsertBefore(_value)
				
				Dim wdSErrors As W.ProofreadingErrors = Document.SpellingErrors
				
				Dim sErrors As New List(Of SpellingError)
				
				If Not wdSErrors Is Nothing Then
				
					For Each wdSError As W.Range In wdSErrors
					
						Dim keepError As Boolean = True
						
						If Not _ignoreWords Is Nothing Then
						
							For j As Integer = 0 To _ignoreWords.Length - 1
							
								If String.Compare(wdSError.Text, _ignoreWords(j), True) = 0 Then
								
									keepError = False
									Exit For
									
								End If
								
							Next
							
						End If
						
						If keepError Then
						
							Dim sError As New SpellingError(wdSError.Text, wdSError.Start)
							
							Dim expandedNumbers As Integer = 5
							
							Dim wdPrevious As W.Range = wdSError
							Dim wdNext As W.Range = wdSError
							
							For i As Integer = 1 To expandedNumbers
							
								If Not wdPrevious Is Nothing Then
								
									wdPrevious = wdPrevious.Previous(W.WdUnits.wdWord, 1)
									
									If Not wdPrevious Is Nothing Then sError.PreviousWords.Insert(0, wdPrevious.Text.Trim)
									
								End If
								
								If Not wdNext Is Nothing Then
								
									wdNext = wdNext.Next(W.WdUnits.wdWord, 1)
									
									If Not wdNext Is Nothing Then sError.NextWords.Add(wdNext.Text.Trim)
									
								End If
								
							Next
							
							If Not wdPrevious Is Nothing Then Marshal.ReleaseComObject(wdPrevious)
							wdPrevious = Nothing
							
							If Not wdNext Is Nothing Then Marshal.ReleaseComObject(wdNext)
							wdNext = Nothing
							
							Dim wdSuggestions As W.SpellingSuggestions = wdApp.GetSpellingSuggestions(wdSError.Text)
							
							If Not wdSuggestions Is Nothing AndAlso wdSuggestions.Count > 0 Then
							
								For Each wdSuggestion As W.SpellingSuggestion In wdSuggestions
								
									sError.Suggestions.Add(New SpellingSuggestion(wdSuggestion.Name))
									
									Marshal.ReleaseComObject(wdSuggestion)
									wdSuggestion = Nothing
									
								Next
								
								Marshal.ReleaseComObject(wdSuggestions)
								wdSuggestions = Nothing
								
							End If
							
							sErrors.Add(sError)
							
						End If
						
						Marshal.ReleaseComObject(wdSError)
						wdSError = Nothing
						
					Next
					
				End If
				
				Marshal.ReleaseComObject(wdSErrors)
				wdSErrors = Nothing
				
				Marshal.ReleaseComObject(wdWord)
				wdWord = Nothing
				
				Marshal.ReleaseComObject(wdWords)
				wdWords = Nothing
				
				Marshal.ReleaseComObject(wdApp)
				wdApp = Nothing
				
				Dim wdRange As W.Range = Document.Range(0, Document.Characters.Count - 1)
				wdRange.Delete()
				
				Marshal.ReleaseComObject(wdRange)
				wdRange = Nothing
				
				Return sErrors.ToArray
				
			End Function
			
			Public Function GetGrammarErrors( _
				ByVal _value As String, _
				ByVal ParamArray _ignoreWords As String() _
			) As GrammaticalError()
			
				Dim grammarValue As Integer = 0
				
				Dim wdApp As W.Application = Document.Application
				Dim wdWords As W.Words = Document.Words
				Dim wdWord As W.Range = wdWords.First
				
				wdWord.InsertBefore(_value)
				
				Dim wdGErrors As W.ProofreadingErrors = Document.GrammaticalErrors
				
				Dim gErrors As New List(Of GrammaticalError)
				
				If Not wdGErrors Is Nothing Then
				
					For Each wdGError As W.Range In wdGErrors
					
						gErrors.Add(New GrammaticalError(wdGError.Text, wdGError.Start))
						
						Marshal.ReleaseComObject(wdGError)
						wdGError = Nothing
						
					Next
					
				End If
				
				Marshal.ReleaseComObject(wdGErrors)
				wdGErrors = Nothing
				
				Marshal.ReleaseComObject(wdWord)
				wdWord = Nothing
				
				Marshal.ReleaseComObject(wdWords)
				wdWords = Nothing
				
				Marshal.ReleaseComObject(wdApp)
				wdApp = Nothing
				
				Dim wdRange As W.Range = Document.Range(0, Document.Characters.Count - 1)
				wdRange.Delete()
				
				Marshal.ReleaseComObject(wdRange)
				wdRange = Nothing
				
				Return gErrors.ToArray
				
			End Function
			
			Public Sub GetErrors( _
				ByVal _value As String, _
				ByRef _spellingErrors As SpellingError(), _
				ByRef _grammarErrors As GrammaticalError(), _
				ByVal ParamArray _ignoreWords As String() _
			)
			
				Dim spellingValue As Integer = 0
				Dim grammarValue As Integer = 0
				
				Dim wdApp As W.Application = Document.Application
				Dim wdWords As W.Words = Document.Words
				Dim wdWord As W.Range = wdWords.First
				
				wdWord.InsertBefore(_value)
				
				Dim wdSErrors As W.ProofreadingErrors = Document.SpellingErrors
				Dim wdGErrors As W.ProofreadingErrors = Document.GrammaticalErrors
				
				Dim sErrors As New List(Of SpellingError)
				
				If Not wdSErrors Is Nothing Then
				
					For Each wdSError As W.Range In wdSErrors
					
						Dim keepError As Boolean = True
						
						If Not _ignoreWords Is Nothing Then
						
							For j As Integer = 0 To _ignoreWords.Length - 1
							
								If String.Compare(wdSError.Text, _ignoreWords(j), True) = 0 Then
								
									keepError = False
									Exit For
									
								End If
								
							Next
							
						End If
						
						If keepError Then
						
							Dim sError As New SpellingError(wdSError.Text, wdSError.Start)
							
							Dim expandedNumbers As Integer = 5
							
							Dim wdPrevious As W.Range = wdSError
							Dim wdNext As W.Range = wdSError
							
							For i As Integer = 1 To expandedNumbers
							
								If Not wdPrevious Is Nothing Then
								
									wdPrevious = wdPrevious.Previous(W.WdUnits.wdWord, 1)
									
									If Not wdPrevious Is Nothing Then sError.PreviousWords.Insert(0, wdPrevious.Text.Trim)
									
								End If
								
								If Not wdNext Is Nothing Then
								
									wdNext = wdNext.Next(W.WdUnits.wdWord, 1)
									
									If Not wdNext Is Nothing Then sError.NextWords.Add(wdNext.Text.Trim)
									
								End If
								
							Next
							
							If Not wdPrevious Is Nothing Then Marshal.ReleaseComObject(wdPrevious)
							wdPrevious = Nothing
							
							If Not wdNext Is Nothing Then Marshal.ReleaseComObject(wdNext)
							wdNext = Nothing
							
							Dim wdSuggestions As W.SpellingSuggestions = wdApp.GetSpellingSuggestions(wdSError.Text)
							
							If Not wdSuggestions Is Nothing AndAlso wdSuggestions.Count > 0 Then
							
								For Each wdSuggestion As W.SpellingSuggestion In wdSuggestions
								
									sError.Suggestions.Add(New SpellingSuggestion(wdSuggestion.Name))
									
									Marshal.ReleaseComObject(wdSuggestion)
									wdSuggestion = Nothing
									
								Next
								
								Marshal.ReleaseComObject(wdSuggestions)
								wdSuggestions = Nothing
								
							End If
							
							sErrors.Add(sError)
							
						End If
						
						Marshal.ReleaseComObject(wdSError)
						wdSError = Nothing
						
					Next
					
				End If
				
				Dim gErrors As New List(Of GrammaticalError)
				
				If Not wdGErrors Is Nothing Then
				
					For Each wdGError As W.Range In wdGErrors
					
						gErrors.Add(New GrammaticalError(wdGError.Text, wdGError.Start))
						
						Marshal.ReleaseComObject(wdGError)
						wdGError = Nothing
						
					Next
					
				End If
				
				Marshal.ReleaseComObject(wdSErrors)
				wdSErrors = Nothing
				
				Marshal.ReleaseComObject(wdGErrors)
				wdGErrors = Nothing
				
				Marshal.ReleaseComObject(wdWord)
				wdWord = Nothing
				
				Marshal.ReleaseComObject(wdWords)
				wdWords = Nothing
				
				Marshal.ReleaseComObject(wdApp)
				wdApp = Nothing
				
				Dim wdRange As W.Range = Document.Range(0, Document.Characters.Count - 1)
				wdRange.Delete()
				
				Marshal.ReleaseComObject(wdRange)
				wdRange = Nothing
				
				_spellingErrors = sErrors.ToArray
				_grammarErrors = gErrors.ToArray
				
			End Sub
			
			Public Function CheckSpelling( _
				ByRef value As String _
			) As Boolean
			
				Dim retVal As Boolean = False
				
				Dim wdApp As W.Application = Document.Application
				
				If Not Document.Application.Visible Then
				
					' Suppress Messages?
					
				End If
				
				Document.Words.First.InsertBefore(value)
				
				Dim wdSpellingErrors As W.ProofreadingErrors = Document.SpellingErrors
				Dim spellingErrorCount As Integer = wdSpellingErrors.Count
				
				For Each wdSpellingError As W.Range In wdSpellingErrors
				
					Dim word As String = wdSpellingError.Text
					Dim wdSuggestions As W.SpellingSuggestions = wdApp.GetSpellingSuggestions(word)
					
					For Each wdSuggestion As W.SpellingSuggestion In wdSuggestions
					
						Dim suggestion As String = wdSuggestion.Name
						
						wdSpellingError.Text = suggestion
						
					Next
					
				Next
				
				' Document.CheckSpelling()
				
				Marshal.ReleaseComObject(wdApp)
				wdApp = Nothing
				
				Dim wdRange As W.Range = Document.Range(0, Document.Characters.Count - 1)
				
				If spellingErrorCount > 0 Then
				
					value = wdRange.Text
					
					retVal = True
					
				End If
				
				wdRange.Delete()
				
				Marshal.ReleaseComObject(wdRange)
				wdRange = Nothing
				
				Return retVal
				
			End Function
			
			Public Function CheckGrammar( _
				ByRef value As String _
			) As Boolean
			
				Document.Words.First.InsertBefore(value)
				
				Dim wdGrammarErrors As W.ProofreadingErrors = Document.GrammaticalErrors
				Dim grammarErrorCount As Integer = wdGrammarErrors.Count
				
				For Each wdGrammarError As W.Range In wdGrammarErrors
				
					Dim word As String = wdGrammarError.Text
					
					'Dim wdSuggestions As W.SpellingSuggestions = wdApp.GetSpellingSuggestions(word)
					
					'For Each wdSuggestion As W.SpellingSuggestion In wdSuggestions
					
						'Dim suggestion As String = wdSuggestion.Name
						
						'wdSpellingError.Text = suggestion
						
					'Next
					
				Next
				
				Dim errors As W.ProofreadingErrors = Document.GrammaticalErrors
				
				If errors.Count > 0 Then
				
					value = Document.Range(0, Document.Characters.Count - 1).Text
					
					Return True
					
				End If
				
				Return False
				
			End Function
			
		#End Region
		
		#Region " Public Shared Methods "
		
			Public Shared Function TryParse( _
				ByVal value As String, _
				ByRef result As DocumentWrapper _
			) As Boolean
			
				Try
				
					result = New DocumentWrapper().ParseFromString(value, A.Word)
					
					Return True
					
				Catch ex As Exception
				End Try
				
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
				
					If Not m_Document Is Nothing Then
					
						If State <> OfficeDocumentState.Existing Then
						
							' Get a reference to the Application
							Dim l_Application As W.Application = m_Document.Application
							
							m_Document.Close(SaveChanges:=False)
							
							If OwnsApplicationInstance Then
							
								Dim l_Documents As W.Documents = l_Application.Documents
								
								If l_Documents.Count = 0 Then
								
									' Kill Documents Reference
									Marshal.ReleaseComObject(l_Documents)
									l_Documents = Nothing
									
									' Close Application
									l_Application.DisplayAlerts = False
									l_Application.Quit()
									
								End If
								
							End If
							
							' Kill Application Reference.
							Marshal.ReleaseComObject(l_Application)
							l_Application = Nothing
							
						End If
						
						' Kill Drawing Reference
						Marshal.ReleaseComObject(m_Document)
						m_Document = Nothing
						
					End If
					
				End If
				
				disposedValue = True
				
			End Sub
			
		#End Region
		
	End Class
	
End Namespace