#tag Module
Protected Module CodeCleanWindowHelpers
	#tag Method, Flags = &h21
		Private Function BuildCodeSmellsSection(analyzer As ProjectAnalyzer) As String
		  // Function BuildCodeSmellsSection(analyzer As ProjectAnalyzer) As String
		  ' Build code smells section for text reports
		  
		  Var report As String = ""
		  
		  report = report + EndOfLine + "CODE SMELL DETECTION" + EndOfLine
		  report = report + "====================" + EndOfLine + EndOfLine
		  
		  Var smells() As CodeSmell = analyzer.DetectCodeSmells()
		  
		  If smells.Count = 0 Then
		    report = report + "✓ No code smells detected - excellent code quality!" + EndOfLine
		    Return report
		  End If
		  
		  ' Group by type
		  Var godClass() As CodeSmell
		  Var featureEnvy() As CodeSmell
		  Var magicNumbers() As CodeSmell
		  Var longParams() As CodeSmell
		  Var deepNesting() As CodeSmell
		  Var deadCode() As CodeSmell
		  Var shotgunSurgery() As CodeSmell
		  
		  For Each smell As CodeSmell In smells
		    Select Case smell.SmellType
		    Case "God Class"
		      godClass.Add(smell)
		    Case "Feature Envy"
		      featureEnvy.Add(smell)
		    Case "Magic Numbers"
		      magicNumbers.Add(smell)
		    Case "Long Parameter List"
		      longParams.Add(smell)
		    Case "Deep Nesting"
		      deepNesting.Add(smell)
		    Case "Dead Code"
		      deadCode.Add(smell)
		    Case "Shotgun Surgery Risk"
		      shotgunSurgery.Add(smell)
		    End Select
		  Next
		  
		  ' Summary
		  report = report + "Total Code Smells Found: " + smells.Count.ToString + EndOfLine + EndOfLine
		  
		  report = report + "Breakdown by Type:" + EndOfLine
		  If godClass.Count > 0 Then
		    report = report + "  🔴 God Classes: " + godClass.Count.ToString + EndOfLine
		  End If
		  If featureEnvy.Count > 0 Then
		    report = report + "  🟠 Feature Envy: " + featureEnvy.Count.ToString + EndOfLine
		  End If
		  If magicNumbers.Count > 0 Then
		    report = report + "  🟡 Magic Numbers: " + magicNumbers.Count.ToString + EndOfLine
		  End If
		  If longParams.Count > 0 Then
		    report = report + "  🟠 Long Parameter Lists: " + longParams.Count.ToString + EndOfLine
		  End If
		  If deepNesting.Count > 0 Then
		    report = report + "  🔴 Deep Nesting: " + deepNesting.Count.ToString + EndOfLine
		  End If
		  If deadCode.Count > 0 Then
		    report = report + "  🟡 Dead Code: " + deadCode.Count.ToString + EndOfLine
		  End If
		  If shotgunSurgery.Count > 0 Then
		    report = report + "  🟠 Shotgun Surgery Risks: " + shotgunSurgery.Count.ToString + EndOfLine
		  End If
		  report = report + EndOfLine
		  
		  ' Severity breakdown
		  Var critical As Integer = 0
		  Var high As Integer = 0
		  Var medium As Integer = 0
		  Var low As Integer = 0
		  
		  For Each smell As CodeSmell In smells
		    Select Case smell.Severity
		    Case "CRITICAL"
		      critical = critical + 1
		    Case "HIGH"
		      high = high + 1
		    Case "MEDIUM"
		      medium = medium + 1
		    Case "LOW"
		      low = low + 1
		    End Select
		  Next
		  
		  report = report + "By Severity:" + EndOfLine
		  If critical > 0 Then
		    report = report + "  🔴 CRITICAL: " + critical.ToString + " (fix immediately)" + EndOfLine
		  End If
		  If high > 0 Then
		    report = report + "  🟠 HIGH: " + high.ToString + " (fix soon)" + EndOfLine
		  End If
		  If medium > 0 Then
		    report = report + "  🟡 MEDIUM: " + medium.ToString + " (address when possible)" + EndOfLine
		  End If
		  If low > 0 Then
		    report = report + "  🔵 LOW: " + low.ToString + " (nice to fix)" + EndOfLine
		  End If
		  report = report + EndOfLine
		  
		  ' Detailed listings by severity
		  If critical > 0 Or high > 0 Then
		    report = report + "🔴 CRITICAL & HIGH PRIORITY SMELLS" + EndOfLine
		    report = report + "===================================" + EndOfLine
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "CRITICAL" Or smell.Severity = "HIGH" Then
		        report = report + FormatCodeSmell(smell) + EndOfLine
		      End If
		    Next
		  End If
		  
		  If medium > 0 Then
		    report = report + EndOfLine + "🟡 MEDIUM PRIORITY SMELLS" + EndOfLine
		    report = report + "=========================" + EndOfLine
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "MEDIUM" Then
		        report = report + FormatCodeSmell(smell) + EndOfLine
		      End If
		    Next
		  End If
		  
		  If low > 0 Then
		    report = report + EndOfLine + "🔵 LOW PRIORITY SMELLS" + EndOfLine
		    report = report + "======================" + EndOfLine
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "LOW" Then
		        report = report + FormatCodeSmell(smell) + EndOfLine
		      End If
		    Next
		  End If
		  
		  Return report
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildQualityScoreSection(analyzer As ProjectAnalyzer) As String
		  //Function BuildQualityScoreSection(analyzer As ProjectAnalyzer) As String
		  ' Build the quality score section for text reports
		  
		  Var report As String = ""
		  
		  report = report + EndOfLine + "CODE QUALITY SCORE" + EndOfLine
		  report = report + "==================" + EndOfLine + EndOfLine
		  
		  Var score As QualityScore = analyzer.CalculateQualityScore()
		  
		  ' Overall score with emoji
		  report = report + score.GetScoreEmoji() + " OVERALL SCORE: " + score.OverallScore.ToString("0.0") + "/100"
		  report = report + "  (Grade: " + score.Grade + ")" + EndOfLine + EndOfLine
		  
		  ' Score interpretation
		  If score.OverallScore >= 80 Then
		    report = report + "Status: EXCELLENT - Well-maintained codebase" + EndOfLine
		  ElseIf score.OverallScore >= 60 Then
		    report = report + "Status: GOOD - Minor improvements recommended" + EndOfLine
		  ElseIf score.OverallScore >= 40 Then
		    report = report + "Status: NEEDS WORK - Several issues to address" + EndOfLine
		  Else
		    report = report + "Status: CRITICAL - Immediate attention required" + EndOfLine
		  End If
		  
		  report = report + EndOfLine + "Component Scores (Weighted):" + EndOfLine
		  report = report + "----------------------------" + EndOfLine
		  
		  ' Error Handling (30%)
		  report = report + "Error Handling:  " + score.ErrorHandlingScore.ToString("0.0") + "/100 (Weight: 30%)" + EndOfLine
		  report = report + "  Coverage: " + score.ErrorHandlingCoverage.ToString("0.0") + "% of risky operations protected" + EndOfLine
		  
		  ' Complexity (25%)
		  report = report + "Complexity:      " + score.ComplexityScore.ToString("0.0") + "/100 (Weight: 25%)" + EndOfLine
		  report = report + "  Average: " + score.AverageComplexity.ToString("0.1") + " (Lower is better)" + EndOfLine
		  
		  ' Code Reuse (20%)
		  report = report + "Code Reuse:      " + score.CodeReuseScore.ToString("0.0") + "/100 (Weight: 20%)" + EndOfLine
		  report = report + "  Unused: " + score.UnusedPercentage.ToString("0.1") + "% of code elements" + EndOfLine
		  
		  ' Parameters (15%)
		  report = report + "Parameters:      " + score.ParameterScore.ToString("0.0") + "/100 (Weight: 15%)" + EndOfLine
		  report = report + "  Average: " + score.AverageParameters.ToString("0.1") + " params per method" + EndOfLine
		  
		  ' Documentation (10%)
		  report = report + "Documentation:   " + score.DocumentationScore.ToString("0.0") + "/100 (Weight: 10%)" + EndOfLine
		  report = report + "  Coverage: " + score.DocumentationCoverage.ToString("0.0") + "% of methods documented" + EndOfLine
		  
		  report = report + EndOfLine + "Improvement Recommendations:" + EndOfLine
		  report = report + "----------------------------" + EndOfLine
		  
		  ' Provide specific recommendations based on weak areas
		  If score.ErrorHandlingScore < 70 Then
		    report = report + "⚠️  Add try/catch blocks to risky operations (database, file I/O, network)" + EndOfLine
		  End If
		  
		  If score.ComplexityScore < 70 Then
		    report = report + "⚠️  Refactor complex methods to improve maintainability" + EndOfLine
		  End If
		  
		  If score.CodeReuseScore < 70 Then
		    report = report + "⚠️  Remove unused code to reduce maintenance burden" + EndOfLine
		  End If
		  
		  If score.ParameterScore < 70 Then
		    report = report + "⚠️  Reduce parameter counts using parameter objects or builder pattern" + EndOfLine
		  End If
		  
		  If score.DocumentationScore < 70 Then
		    report = report + "⚠️  Add comments to explain complex logic and method purposes" + EndOfLine
		  End If
		  
		  If score.OverallScore >= 80 Then
		    report = report + "✅ Great work! Continue maintaining these standards." + EndOfLine
		  End If
		  
		  report = report + EndOfLine
		  
		  Return report
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildRefactoringSuggestionsSection(proj As ProjectAnalyzer) As String
		  // Function BuildRefactoringSuggestionsSection(proj As ProjectAnalyzer) As String
		  Var report As String = ""
		  report = report + EndOfLine + "REFACTORING SUGGESTIONS" + EndOfLine
		  report = report + "=====================" + EndOfLine + EndOfLine
		  
		  // Collect and group by priority
		  Var highPriority() As RefactoringSuggestion
		  Var mediumPriority() As RefactoringSuggestion
		  Var lowPriority() As RefactoringSuggestion
		  
		  Var methods() As CodeElement = proj.GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    For Each suggestion As RefactoringSuggestion In method.RefactoringSuggestions
		      Select Case suggestion.Priority
		      Case "HIGH"
		        highPriority.Add(suggestion)
		      Case "MEDIUM"
		        mediumPriority.Add(suggestion)
		      Case "LOW"
		        lowPriority.Add(suggestion)
		      End Select
		    Next
		  Next
		  
		  // DEDUPLICATE - This is the fix!
		  Var generator As New ReportGenerator
		  generator.DeduplicateSuggestionsInPlace(highPriority)
		  generator.DeduplicateSuggestionsInPlace(mediumPriority)
		  generator.DeduplicateSuggestionsInPlace(lowPriority)
		  
		  // Calculate total from deduplicated counts
		  Var totalSuggestions As Integer = highPriority.Count + mediumPriority.Count + lowPriority.Count
		  
		  If totalSuggestions = 0 Then
		    report = report + "✓ No refactoring suggestions - code looks good!" + EndOfLine
		    Return report
		  End If
		  
		  report = report + "Total Suggestions: " + totalSuggestions.ToString + EndOfLine
		  report = report + "  - High Priority: " + highPriority.Count.ToString + EndOfLine
		  report = report + "  - Medium Priority: " + mediumPriority.Count.ToString + EndOfLine
		  report = report + "  - Low Priority: " + lowPriority.Count.ToString + EndOfLine + EndOfLine
		  
		  // Show high priority first
		  If highPriority.Count > 0 Then
		    report = report + "🔴 HIGH PRIORITY (Fix First)" + EndOfLine
		    report = report + "----------------------------" + EndOfLine
		    For Each suggestion As RefactoringSuggestion In highPriority
		      report = report + FormatSuggestion(suggestion) + EndOfLine
		    Next
		  End If
		  
		  // Then medium
		  If mediumPriority.Count > 0 Then
		    report = report + EndOfLine + "🟡 MEDIUM PRIORITY" + EndOfLine
		    report = report + "------------------" + EndOfLine
		    For Each suggestion As RefactoringSuggestion In mediumPriority
		      report = report + FormatSuggestion(suggestion) + EndOfLine
		    Next
		  End If
		  
		  // And low (if any)
		  If lowPriority.Count > 0 Then
		    report = report + EndOfLine + "🔵 LOW PRIORITY" + EndOfLine
		    report = report + "---------------" + EndOfLine
		    For Each suggestion As RefactoringSuggestion In lowPriority
		      report = report + FormatSuggestion(suggestion) + EndOfLine
		    Next
		  End If
		  
		  Return report
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildRelationshipSection(analyzer As ProjectAnalyzer) As String
		  // Generate the relationship analysis section
		  
		  Var output As String = ""
		  
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + "    RELATIONSHIP ANALYSIS" + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + EndOfLine
		  
		  Var methods() As CodeElement = analyzer.GetMethodElements()  // ✅ Call the method
		  
		  
		  Var stats As Dictionary = CalculateRelationshipStats(methods)
		  
		  Var methodsWithCalls As Integer = stats.Value("methodsWithCalls")
		  Var totalCalls As Integer = stats.Value("totalCalls")
		  
		  output = output + "Methods with outgoing calls: " + methodsWithCalls.ToString + EndOfLine
		  output = output + "Total method calls detected: " + totalCalls.ToString + EndOfLine
		  output = output + EndOfLine
		  
		  If totalCalls > 0 Then
		    output = output + BuildSampleRelationships(methods)
		  End If
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildReportFooter() As String
		  // Generate the report footer
		  
		  Var output As String = ""
		  
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + "Scan completed successfully!" + EndOfLine
		  output = output + "Use 'Generate Flowchart PDF' to visualize." + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildReportHeader() As String
		  // Generate the report header
		  
		  Var output As String = ""
		  
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + "    CODE ANALYSIS REPORT" + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + EndOfLine
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildSampleRelationships(methods() As CodeElement) As String
		  // Build a sample list of method call relationships
		  
		  Var output As String = ""
		  
		  output = output + "Sample relationships (first 10):" + EndOfLine
		  
		  Var count As Integer = 0
		  For Each element As CodeElement In methods
		    If element.CallsTo.Count > 0 And count < 10 Then
		      For Each target As CodeElement In element.CallsTo
		        output = output + "  " + element.FullPath + " -> " + target.FullPath + EndOfLine
		        count = count + 1
		        If count >= 10 Then Exit For element
		      Next
		    End If
		  Next
		  
		  output = output + EndOfLine
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildStatisticsSection(analyzer As ProjectAnalyzer) As String
		  // Private Function BuildStatisticsSection(analyzer As ProjectAnalyzer) As String
		  // Generate the statistics summary section
		  Var output As String = ""
		  
		  Var allElements() As CodeElement = analyzer.GetAllElements()      // ✅ CALL METHOD
		  Var classElements() As CodeElement = analyzer.GetClassElements()   // ✅ CALL METHOD
		  Var moduleElements() As CodeElement = analyzer.GetModuleElements() // ✅ CALL METHOD
		  Var methodElements() As CodeElement = analyzer.GetMethodElements() // ✅ CALL METHOD
		  
		  output = output + "Total Elements Found: " + allElements.Count.ToString + EndOfLine
		  output = output + "  - Classes: " + classElements.Count.ToString + EndOfLine
		  output = output + "  - Modules: " + moduleElements.Count.ToString + EndOfLine
		  output = output + "  - Methods: " + methodElements.Count.ToString + EndOfLine
		  output = output + EndOfLine
		  
		  Return output
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildUnusedByTypeList(unused() As CodeElement) As String
		  // Group unused elements by type and format for display
		  
		  Var output As String = ""
		  
		  // Group by type
		  Var unusedByType As Dictionary = GroupElementsByType(unused)
		  
		  // Display each type
		  For Each key As Variant In unusedByType.Keys
		    Var typeName As String = key.StringValue
		    Var elements() As CodeElement = unusedByType.Value(key)
		    
		    output = output + "▼ " + typeName + " (" + elements.Count.ToString + "):" + EndOfLine
		    output = output + FormatElementList(elements)
		    output = output + EndOfLine
		  Next
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildUnusedElementsSection(analyzer As ProjectAnalyzer) As String
		  // Generate the unused elements section
		  
		  Var output As String = ""
		  
		  Var unused() As CodeElement = analyzer.GetUnusedElements()
		  
		  If unused.Count = 0 Then
		    output = output + "✓ Great! No unused code elements found." + EndOfLine
		  Else
		    output = output + "⚠ Found " + unused.Count.ToString + " UNUSED elements:" + EndOfLine
		    output = output + EndOfLine
		    output = output + BuildUnusedByTypeList(unused)
		  End If
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateRelationshipStats(methods() As CodeElement) As Dictionary
		  // Calculate statistics about method relationships
		  
		  Var stats As New Dictionary
		  Var methodsWithCalls As Integer = 0
		  Var totalCalls As Integer = 0
		  
		  For Each element As CodeElement In methods
		    If element.CallsTo.Count > 0 Then
		      methodsWithCalls = methodsWithCalls + 1
		      totalCalls = totalCalls + element.CallsTo.Count
		    End If
		  Next
		  
		  stats.Value("methodsWithCalls") = methodsWithCalls
		  stats.Value("totalCalls") = totalCalls
		  
		  Return stats
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FormatCodeSmell(smell As CodeSmell) As String
		  // Private Function FormatCodeSmell(smell As CodeSmell) As String
		  ' Format a single code smell for text display
		  
		  Var output As String = ""
		  
		  output = output + smell.GetSeverityEmoji() + " " + smell.SmellType + ": " + smell.Element.FullPath + EndOfLine
		  output = output + "   " + smell.Description + EndOfLine
		  output = output + "   " + smell.Details + EndOfLine
		  output = output + "   → " + smell.Recommendation + EndOfLine
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FormatElementList(elements() As CodeElement) As String
		  // Format a list of elements for display
		  
		  Var output As String = ""
		  
		  For Each element As CodeElement In elements
		    output = output + "  • " + element.FullPath
		    If element.FileName.Trim <> "" Then
		      output = output + " [" + element.FileName + "]"
		    End If
		    output = output + EndOfLine
		  Next
		  
		  Return output
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FormatSuggestion(suggestion As RefactoringSuggestion) As String
		  // Private Function FormatSuggestion(suggestion As RefactoringSuggestion) As String
		  Var output As String = ""
		  
		  output = output + "📍 " + suggestion.Element.FullPath + EndOfLine
		  output = output + "   " + suggestion.Title + EndOfLine
		  output = output + "   " + suggestion.Description + EndOfLine
		  output = output + EndOfLine
		  
		  For Each tip As String In suggestion.Suggestions
		    If tip.Trim <> "" Then
		      output = output + "   → " + tip + EndOfLine
		    Else
		      output = output + EndOfLine
		    End If
		  Next
		  
		  output = output + EndOfLine
		  
		  Return output
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GenerateTextReport(Extends win As CodeCleanWindow)
		  //Function GenerateTextReport(Extends win As CodeCleanWindow)
		  
		  // Function GenerateTextReport(Extends win As CodeCleanWindow)
		  If win.mAnalyzer = Nil Then Return
		  
		  Var output As String = ""
		  
		  output = output + BuildReportHeader()
		  output = output + BuildQualityScoreSection(win.mAnalyzer)
		  output = output + BuildCodeSmellsSection(win.mAnalyzer)  // ← ADD THIS
		  output = output + BuildStatisticsSection(win.mAnalyzer)
		  output = output + BuildUnusedElementsSection(win.mAnalyzer)
		  output = output + BuildRelationshipSection(win.mAnalyzer)
		  output = output + BuildRefactoringSuggestionsSection(win.mAnalyzer)
		  output = output + BuildReportFooter()
		  
		  win.txtResults.Text = output
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GroupElementsByType(elements() As CodeElement) As Dictionary
		  // Group elements into a dictionary by their type
		  
		  Var grouped As New Dictionary
		  
		  For Each element As CodeElement In elements
		    If Not grouped.HasKey(element.ElementType) Then
		      Var arr() As CodeElement
		      grouped.Value(element.ElementType) = arr
		    End If
		    
		    Var arr() As CodeElement = grouped.Value(element.ElementType)
		    arr.Add(element)
		    grouped.Value(element.ElementType) = arr
		  Next
		  
		  Return grouped
		End Function
	#tag EndMethod


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
