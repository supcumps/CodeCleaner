#tag Module
Protected Module HotSpotsGenerator
	#tag Method, Flags = &h21
		Private Function AnalyzeMethodForHotSpot(element As CodeElement) As HotSpot
		  //Private Function AnalyzeMethodForHotSpot(element As CodeElement) As HotSpot
		  Var hs As New HotSpot
		  
		  hs.MethodPath = element.FullPath
		  hs.ComplexityScore = element.CyclomaticComplexity
		  
		  // DEBUG: Check what we're getting
		  System.DebugLog("Method: " + element.FullPath + " - Complexity: " + element.CyclomaticComplexity.ToString)
		  
		  hs.ParameterCount = element.ParameterCount
		  hs.LinesOfCode = element.LinesOfCode
		  
		  
		  
		  // Get refactoring suggestions
		  Var suggestions() As RefactoringSuggestion = element.RefactoringSuggestions()
		  
		  // Count issues from refactoring suggestions
		  Var issueTypes() As String
		  
		  For i As Integer = 0 To suggestions.LastIndex
		    Var suggestion As RefactoringSuggestion = suggestions(i)
		    hs.IssueCount = hs.IssueCount + 1
		    
		    Var categoryExists As Boolean = False
		    For j As Integer = 0 To issueTypes.LastIndex
		      If issueTypes(j) = suggestion.Category Then
		        categoryExists = True
		        Exit For j
		      End If
		    Next j
		    
		    If Not categoryExists Then
		      issueTypes.Add(suggestion.Category)
		    End If
		  Next i
		  
		  hs.IssueTypes = issueTypes
		  
		  // Check for missing error handling
		  hs.MissingErrorHandling = CheckMissingErrorHandling(element)
		  
		  // Store CalledBy count
		  hs.CalledByCount = element.CalledBy.Count
		  
		  // Calculate hot spot score
		  hs.HotSpotScore = CalculateHotSpotScore(hs, element)
		  
		  // Determine risk level
		  hs.RiskLevel = DetermineRiskLevel(hs)
		  
		  // Generate impact description
		  hs.ImpactDescription = GenerateImpactDescription(hs, element)
		  
		  Return hs
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateHotSpotScore(hs As HotSpot, element As CodeElement) As Integer
		  //Private Function CalculateHotSpotScore(hs As HotSpot, element As CodeElement) As Integer
		  Var score As Integer = 0
		  
		  // Complexity contributes heavily
		  If hs.ComplexityScore >= 20 Then
		    score = score + 30
		  ElseIf hs.ComplexityScore >= 15 Then
		    score = score + 20
		  ElseIf hs.ComplexityScore >= 10 Then
		    score = score + 10
		  End If
		  
		  // Deep nesting is a major issue
		  If hs.NestingDepth >= 6 Then
		    score = score + 25
		  ElseIf hs.NestingDepth >= 4 Then
		    score = score + 15
		  End If
		  
		  // Too many parameters
		  If hs.ParameterCount >= 6 Then
		    score = score + 20
		  ElseIf hs.ParameterCount >= 4 Then
		    score = score + 10
		  End If
		  
		  // Long methods
		  If hs.LinesOfCode >= 200 Then
		    score = score + 20
		  ElseIf hs.LinesOfCode >= 100 Then
		    score = score + 10
		  End If
		  
		  // Missing error handling in complex code
		  If hs.MissingErrorHandling And hs.ComplexityScore >= 10 Then
		    score = score + 25
		  ElseIf hs.MissingErrorHandling Then
		    score = score + 10
		  End If
		  
		  // Multiple issue types compound the problem
		  score = score + (hs.IssueTypes.Count * 5)
		  
		  // Methods that are called frequently (CalledBy count)
		  If hs.CalledByCount >= 10 Then
		    score = score + 20
		  ElseIf hs.CalledByCount >= 5 Then
		    score = score + 10
		  End If
		  
		  Return score
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CheckMissingErrorHandling(element As CodeElement) As Boolean
		  // Private Function CheckMissingErrorHandling(element As CodeElement) As Boolean
		  // Get the suggestions array first
		  Var suggestions() As RefactoringSuggestion = element.RefactoringSuggestions()
		  
		  // Iterate through using index-based loop
		  For i As Integer = 0 To suggestions.LastIndex
		    Var suggestion As RefactoringSuggestion = suggestions(i)
		    If suggestion.Category = "Missing Error Handling" Then
		      Return True
		    End If
		  Next i
		  
		  Return False
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CompareHotSpots(hs1 As HotSpot, hs2 As HotSpot) As Integer
		  // Private Function CompareHotSpots(hs1 As HotSpot, hs2 As HotSpot) As Integer
		  // Sort by score descending (highest first)
		  If hs1.HotSpotScore > hs2.HotSpotScore Then
		    Return -1
		  ElseIf hs1.HotSpotScore < hs2.HotSpotScore Then
		    Return 1
		  Else
		    Return 0
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetermineRiskLevel(hs As HotSpot) As String
		  // Private Function DetermineRiskLevel(hs As HotSpot) As String
		  If hs.HotSpotScore >= 80 Then
		    Return "CRITICAL"
		  ElseIf hs.HotSpotScore >= 50 Then
		    Return "HIGH"
		  ElseIf hs.HotSpotScore >= 25 Then
		    Return "MEDIUM"
		  Else
		    Return "LOW"
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GenerateHotSpots(elements() As CodeElement) As HotSpot()
		  //Public Function GenerateHotSpots(elements() As CodeElement) As HotSpot()
		  Var hotSpots() As HotSpot
		  
		  For i As Integer = 0 To elements.LastIndex
		    Var element As CodeElement = elements(i)
		    If element.ElementType = "Method" Then
		      Var hotSpot As HotSpot = AnalyzeMethodForHotSpot(element)
		      If hotSpot <> Nil And hotSpot.HotSpotScore >= 10 Then
		        hotSpots.Add(hotSpot)
		      End If
		    End If
		  Next
		  
		  // Sort by hot spot score (highest first)
		  If hotSpots.LastIndex >= 0 Then
		    SortHotSpotsByScore(hotSpots)
		  End If
		  
		  Return hotSpots
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GenerateImpactDescription(hs As HotSpot, element As CodeElement) As String
		  // Private Function GenerateImpactDescription(hs As HotSpot, element As CodeElement) As String
		  Var descriptions() As String
		  
		  If hs.ComplexityScore >= 15 Then
		    descriptions.Add("Very difficult to understand and modify")
		  End If
		  
		  If hs.NestingDepth >= 6 Then
		    descriptions.Add("Deeply nested logic makes debugging challenging")
		  End If
		  
		  If hs.MissingErrorHandling Then
		    descriptions.Add("Lacks error handling - prone to crashes")
		  End If
		  
		  If hs.LinesOfCode >= 150 Then
		    descriptions.Add("Excessively long - violates single responsibility principle")
		  End If
		  
		  If hs.CalledByCount >= 10 Then
		    descriptions.Add("Called by " + hs.CalledByCount.ToString + " other methods - changes have wide impact")
		  End If
		  
		  If hs.ParameterCount >= 6 Then
		    descriptions.Add("Too many parameters indicate poor abstraction")
		  End If
		  
		  If descriptions.Count = 0 Then
		    Return "Multiple moderate quality issues"
		  Else
		    Return String.FromArray(descriptions, "; ")
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetHotSpotsByRiskLevel(hotSpots() As HotSpot, riskLevel As String) As HotSpot()
		  // Public Function GetHotSpotsByRiskLevel(hotSpots() As HotSpot, riskLevel As String) As HotSpot()
		  Var filtered() As HotSpot
		  
		  For i As Integer = 0 To hotSpots.LastIndex
		    If hotSpots(i).RiskLevel = riskLevel Then
		      filtered.Add(hotSpots(i))
		    End If
		  Next
		  
		  Return filtered
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTopHotSpots(hotSpots() As HotSpot, count As Integer) As HotSpot()
		  //Public Function GetTopHotSpots(hotSpots() As HotSpot, count As Integer) As HotSpot()
		  Var result() As HotSpot
		  Var limit As Integer
		  
		  If count < (hotSpots.LastIndex + 1) Then
		    limit = count
		  Else
		    limit = hotSpots.LastIndex + 1
		  End If
		  
		  For i As Integer = 0 To limit - 1
		    result.Add(hotSpots(i))
		  Next
		  
		  Return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HasMissingErrorHandling(element As CodeElement) As Boolean
		  // Private Function HasMissingErrorHandling(element As CodeElement) As Boolean
		  // Get the suggestions array first
		  Var suggestions() As RefactoringSuggestion = element.RefactoringSuggestions()
		  
		  // Check if array has any elements
		  If suggestions.LastIndex < 0 Then
		    Return False
		  End If
		  
		  // Iterate through using index-based loop (not For Each)
		  For i As Integer = 0 To suggestions.LastIndex
		    Var suggestion As RefactoringSuggestion = suggestions(i)
		    If suggestion.Category = "Missing Error Handling" Then
		      Return True
		    End If
		  Next i
		  
		  Return False
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SortHotSpotsByScore(hotSpots() As HotSpot)
		  //Private Sub SortHotSpotsByScore(hotSpots() As HotSpot)
		  // Simple bubble sort for hot spots by score (descending)
		  For i As Integer = 0 To hotSpots.LastIndex
		    For j As Integer = i + 1 To hotSpots.LastIndex
		      If hotSpots(j).HotSpotScore > hotSpots(i).HotSpotScore Then
		        // Swap
		        Var temp As HotSpot = hotSpots(i)
		        hotSpots(i) = hotSpots(j)
		        hotSpots(j) = temp
		      End If
		    Next j
		  Next i
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = Purpose
		 code elements and generates hot spot report
		
	#tag EndNote


	#tag Property, Flags = &h0
		Untitled As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
