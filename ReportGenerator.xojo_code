#tag Class
Protected Class ReportGenerator
	#tag Method, Flags = &h21
		Private Function CalculateAggregateStats(methodMetrics() As MethodMetrics) As Dictionary
		  // Calculate aggregate statistics from method metrics
		  
		  Var stats As New Dictionary
		  
		  Var totalLOC As Integer = 0
		  Var totalComplexity As Integer = 0
		  Var maxComplexity As Integer = 0
		  Var maxLOC As Integer = 0
		  Var maxCallDepth As Integer = 0
		  Var maxChainDepth As Integer = 0
		  Var maxParameterCount As Integer = 0
		  Var methodsWithTooManyParams As Integer = 0
		  Var totalParameters As Integer = 0
		  Var totalOptionalParams As Integer = 0
		  
		  For Each metric As MethodMetrics In methodMetrics
		    totalLOC = totalLOC + metric.LinesOfCode
		    totalComplexity = totalComplexity + metric.Complexity
		    totalParameters = totalParameters + metric.ParameterCount
		    totalOptionalParams = totalOptionalParams + metric.OptionalParameterCount
		    
		    If metric.Complexity > maxComplexity Then
		      maxComplexity = metric.Complexity
		    End If
		    If metric.LinesOfCode > maxLOC Then
		      maxLOC = metric.LinesOfCode
		    End If
		    If metric.CallDepth > maxCallDepth Then
		      maxCallDepth = metric.CallDepth
		    End If
		    If metric.CallChainDepth > maxChainDepth Then
		      maxChainDepth = metric.CallChainDepth
		    End If
		    If metric.ParameterCount > maxParameterCount Then
		      maxParameterCount = metric.ParameterCount
		    End If
		    If metric.HasTooManyParameters Then
		      methodsWithTooManyParams = methodsWithTooManyParams + 1
		    End If
		  Next
		  
		  Var avgLOC As Integer = 0
		  Var avgComplexity As Integer = 0
		  Var avgParameters As Integer = 0
		  If methodMetrics.Count > 0 Then
		    avgLOC = totalLOC \ methodMetrics.Count
		    avgComplexity = totalComplexity \ methodMetrics.Count
		    avgParameters = totalParameters \ methodMetrics.Count
		  End If
		  
		  stats.Value("totalLOC") = totalLOC
		  stats.Value("avgLOC") = avgLOC
		  stats.Value("avgComplexity") = avgComplexity
		  stats.Value("maxComplexity") = maxComplexity
		  stats.Value("maxLOC") = maxLOC
		  stats.Value("maxChainDepth") = maxChainDepth
		  stats.Value("maxParameterCount") = maxParameterCount
		  stats.Value("avgParameters") = avgParameters
		  stats.Value("methodsWithTooManyParams") = methodsWithTooManyParams
		  stats.Value("totalOptionalParams") = totalOptionalParams
		  
		  Return stats
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateCallChainDepth(element As CodeElement, visited As Dictionary, currentDepth As Integer) As Integer
		  // Calculate the deepest call chain from this element
		  
		  If visited.HasKey(element.FullPath) Then
		    Return currentDepth
		  End If
		  
		  visited.Value(element.FullPath) = True
		  
		  Var maxDepth As Integer = currentDepth
		  
		  For Each calledElement As CodeElement In element.CallsTo
		    Var depth As Integer = CalculateCallChainDepth(calledElement, visited, currentDepth + 1)
		    If depth > maxDepth Then
		      maxDepth = depth
		    End If
		  Next
		  
		  visited.Remove(element.FullPath)
		  
		  Return maxDepth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateCyclomaticComplexity(code As String) As Integer
		  // Count decision points
		  
		  Var complexity As Integer = 1
		  Var upperCode As String = code.Uppercase
		  
		  complexity = complexity + CountOccurrences(upperCode, " IF ")
		  complexity = complexity + CountOccurrences(upperCode, EndOfLine + "IF ")
		  complexity = complexity + CountOccurrences(upperCode, "ELSEIF ")
		  complexity = complexity + CountOccurrences(upperCode, " FOR ")
		  complexity = complexity + CountOccurrences(upperCode, EndOfLine + "FOR ")
		  complexity = complexity + CountOccurrences(upperCode, " WHILE ")
		  complexity = complexity + CountOccurrences(upperCode, EndOfLine + "WHILE ")
		  complexity = complexity + CountOccurrences(upperCode, " DO ")
		  complexity = complexity + CountOccurrences(upperCode, EndOfLine + "DO ")
		  complexity = complexity + CountOccurrences(upperCode, " CASE ")
		  complexity = complexity + CountOccurrences(upperCode, EndOfLine + "CASE ")
		  complexity = complexity + CountOccurrences(upperCode, " AND ")
		  complexity = complexity + CountOccurrences(upperCode, " OR ")
		  
		  Return complexity
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculatePDFHeight(analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double) As Integer
		  // Private Function CalculatePDFHeight(analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double) As Integer
		  ' More conservative height calculation
		  
		  Var smells() As CodeSmell = analyzer.DetectCodeSmells()
		  Var stats As Dictionary = analyzer.GetErrorHandlingStats()
		  Var riskyCount As Integer = 0
		  
		  If stats.HasKey("riskyMethods") Then
		    Var riskyMethods() As CodeElement = stats.Value("riskyMethods")
		    riskyCount = riskyMethods.Count
		  End If
		  
		  // REDUCED estimates per item
		  Var estimatedHeight As Integer = 2500  // Base sections (reduced from 3000)
		  estimatedHeight = estimatedHeight + (smells.Count * 100)  // 100 pixels per smell (reduced from 150)
		  estimatedHeight = estimatedHeight + (riskyCount * 80)  // 80 pixels per method (reduced from 100)
		  
		  // Add 20% buffer
		  estimatedHeight = Round(estimatedHeight * 1.2)
		  
		  // Bounds
		  If estimatedHeight < 5000 Then estimatedHeight = 5000
		  If estimatedHeight > 18000 Then estimatedHeight = 18000  // Even lower max
		  
		  System.DebugLog("=== PDF HEIGHT CALCULATION ===")
		  System.DebugLog("Code Smells: " + str(smells.Count)+ " × 100 = " + Str(smells.Count * 100) )
		  System.DebugLog("Risky Methods: " + Str(riskyCount) + " × 80 = " + Str(riskyCount * 80))
		  System.DebugLog("FINAL HEIGHT: " + estimatedHeight.ToString + " pixels")
		  
		  Return estimatedHeight
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateRefactoringPDFHeight(analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double) As Integer
		  // Private Function CalculateRefactoringPDFHeight(analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double) As Integer
		  Var height As Double = margin * 3
		  
		  // Header
		  height = height + 120
		  
		  // Summary section
		  height = height + 150
		  
		  // Collect all suggestions and DEDUPLICATE
		  Var highPriority() As RefactoringSuggestion
		  Var mediumPriority() As RefactoringSuggestion
		  Var lowPriority() As RefactoringSuggestion
		  
		  Var methods() As CodeElement = analyzer.GetMethodElements()
		  
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
		  
		  // DEDUPLICATE - This is the critical fix!
		  DeduplicateSuggestionsInPlace(highPriority)
		  DeduplicateSuggestionsInPlace(mediumPriority)
		  DeduplicateSuggestionsInPlace(lowPriority)
		  
		  // Calculate space for HIGH priority
		  For Each suggestion As RefactoringSuggestion In highPriority
		    height = height + 40  // Method name and title
		    height = height + 25  // Description
		    height = height + (suggestion.Suggestions.Count * lineHeight)
		    height = height + 30  // Spacing
		  Next
		  
		  // Calculate space for MEDIUM priority
		  For Each suggestion As RefactoringSuggestion In mediumPriority
		    height = height + 40
		    height = height + 25
		    height = height + (suggestion.Suggestions.Count * lineHeight)
		    height = height + 30
		  Next
		  
		  // Calculate space for LOW priority
		  For Each suggestion As RefactoringSuggestion In lowPriority
		    height = height + 40
		    height = height + 25
		    height = height + (suggestion.Suggestions.Count * lineHeight)
		    height = height + 30
		  Next
		  
		  // Footer
		  height = height + 50
		  
		  // Add section headers (3 priorities * 45 pixels each)
		  height = height + 135
		  
		  Return CType(height, Integer)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CountOccurrences(text As String, searchFor As String) As Integer
		  // Count how many times searchFor appears in text
		  
		  Var count As Integer = 0
		  Var pos As Integer = 0
		  
		  Do
		    pos = text.IndexOf(pos, searchFor)
		    If pos >= 0 Then
		      count = count + 1
		      pos = pos + searchFor.Length
		    End If
		  Loop Until pos < 0
		  
		  Return count
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DeduplicateSuggestions(suggestions() As RefactoringSuggestion) As RefactoringSuggestion()
		  // Private Function DeduplicateSuggestions(suggestions() As RefactoringSuggestion) As RefactoringSuggestion()
		  Var uniqueSuggestions() As RefactoringSuggestion
		  Var seenKeys As New Dictionary
		  
		  For Each suggestion As RefactoringSuggestion In suggestions
		    // Create a unique key using Element.Name + Category + Description
		    Var methodName As String = ""
		    If suggestion.Element <> Nil Then
		      methodName = suggestion.Element.Name
		    End If
		    
		    Var key As String = methodName + "|" + suggestion.Category + "|" + suggestion.Description
		    
		    If Not seenKeys.HasKey(key) Then
		      uniqueSuggestions.Add(suggestion)
		      seenKeys.Value(key) = True
		    End If
		  Next
		  
		  Return uniqueSuggestions
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeduplicateSuggestionsInPlace(ByRef suggestions() As RefactoringSuggestion)
		  // Private Sub DeduplicateSuggestionsInPlace(ByRef suggestions() As RefactoringSuggestion)
		  '
		  
		  Var uniqueSuggestions() As RefactoringSuggestion
		  Var seenKeys As New Dictionary
		  
		  For Each suggestion As RefactoringSuggestion In suggestions
		    // Create a unique key
		    Var methodName As String = ""
		    If suggestion.Element <> Nil Then
		      methodName = suggestion.Element.Name
		    End If
		    
		    Var key As String = methodName + "|" + suggestion.Category + "|" + suggestion.Description
		    
		    If Not seenKeys.HasKey(key) Then
		      uniqueSuggestions.Add(suggestion)
		      seenKeys.Value(key) = True
		    End If
		  Next
		  
		  // Clear original array and copy unique ones back
		  suggestions.RemoveAll
		  For Each uniqueSugg As RefactoringSuggestion In uniqueSuggestions
		    suggestions.Add(uniqueSugg)
		  Next
		  
		  
		  
		  
		  '
		  '// Private Sub DeduplicateSuggestionsInPlace(ByRef suggestions() As RefactoringSuggestion)
		  'Var uniqueSuggestions() As RefactoringSuggestion
		  'Var seenKeys As New Dictionary
		  '
		  'For Each suggestion As RefactoringSuggestion In suggestions
		  '// Create a unique key
		  'Var methodName As String = ""
		  'If suggestion.Element <> Nil Then
		  'methodName = suggestion.Element.Name
		  'End If
		  '
		  'Var key As String = methodName + "|" + suggestion.Category + "|" + suggestion.Description
		  '
		  'If Not seenKeys.HasKey(key) Then
		  'uniqueSuggestions.Add(suggestion)
		  'seenKeys.Value(key) = True
		  'End If
		  'Next
		  '
		  '// Replace the original array contents
		  'suggestions.RemoveAll
		  'For Each uniqueSugg As RefactoringSuggestion In uniqueSuggestions
		  'suggestions.Add(uniqueSugg)
		  'Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawCenteredText(g As Graphics, text As String, centerX As Double, y As Double, maxWidth As Double)
		  // Draw text centered at the given x position
		  Var textWidth As Double = g.TextWidth(text)
		  Var x As Double = centerX - (textWidth / 2)
		  g.DrawText(text, x, y)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GenerateAnalysisReportPDF(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean
		  // Function GenerateAnalysisReportPDF(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean
		  Try
		    // Setup
		    Var pageWidth As Integer = 612
		    Var margin As Double = 100
		    Var lineHeight As Double = 14
		    
		    // Calculate height and create PDF
		    Var totalHeight As Integer = CalculatePDFHeight(analyzer, margin, lineHeight)
		    Var pdf As New PDFDocument(pageWidth, totalHeight)
		    Var g As Graphics = pdf.Graphics
		    
		    // White background
		    g.DrawingColor = Color.White
		    g.FillRectangle(0, 0, pageWidth, totalHeight)
		    
		    Var yPos As Double = margin
		    
		    // Render each section
		    yPos = RenderHeader(g, pageWidth, margin, yPos)
		    yPos = RenderQualityScore(g, analyzer, pageWidth, margin, lineHeight, yPos)  
		    yPos = RenderCodeSmells(g, analyzer, pageWidth, margin, lineHeight, yPos)
		    yPos = RenderSummary(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderUnusedElements(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderErrorHandlingAnalysis(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderComplexityMetrics(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderParameterComplexityDetails(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderTopComplexMethods(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderRelationships(g, analyzer, margin, lineHeight, yPos)
		    yPos = RenderFooter(g, pageWidth, margin, lineHeight, yPos)
		    
		    // Save
		    pdf.Save(saveFile)
		    Return True
		    
		  Catch e As RuntimeException
		    System.DebugLog("Error generating PDF: " + e.Message)
		    Return False
		  End Try
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GenerateRefactoringSuggestionsReport(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean
		  // Function GenerateRefactoringSuggestionsReport(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean
		  Try
		    System.DebugLog("=== Starting PDF Generation ===")
		    
		    // Setup
		    Var pageWidth As Integer = 612
		    Var margin As Double = 50
		    Var lineHeight As Double = 14
		    
		    System.DebugLog("Calculating height...")
		    // Calculate height needed
		    Var totalHeight As Integer = CalculateRefactoringPDFHeight(analyzer, margin, lineHeight)
		    System.DebugLog("Total height: " + totalHeight.ToString)
		    
		    // Add extra height buffer to prevent cutoff
		    totalHeight = totalHeight + 50
		    
		    // Create PDF
		    System.DebugLog("Creating PDF document...")
		    Var pdf As New PDFDocument(pageWidth, totalHeight)
		    Var g As Graphics = pdf.Graphics
		    
		    If g = Nil Then
		      System.DebugLog("ERROR: Graphics context is Nil!")
		      Return False
		    End If
		    System.DebugLog("Graphics context created successfully")
		    
		    // White background from absolute top
		    g.DrawingColor = Color.White
		    g.FillRectangle(0, 0, pageWidth, totalHeight)
		    System.DebugLog("Background filled")
		    
		    Var yPos As Double = margin + 30
		    
		    // Render sections
		    System.DebugLog("Rendering header...")
		    yPos = RenderRefactoringHeader(g, pageWidth, margin, yPos)
		    System.DebugLog("Header rendered. yPos: " + yPos.ToString)
		    
		    System.DebugLog("Rendering summary...")
		    yPos = RenderRefactoringSummary(g, analyzer, margin, lineHeight, yPos)
		    System.DebugLog("Summary rendered. yPos: " + yPos.ToString)
		    
		    System.DebugLog("Rendering suggestions...")
		    yPos = RenderRefactoringSuggestions(g, analyzer, pageWidth, margin, lineHeight, yPos)
		    System.DebugLog("Suggestions rendered. yPos: " + yPos.ToString)
		    
		    System.DebugLog("Rendering footer...")
		    yPos = RenderRefactoringFooter(g, pageWidth, margin, lineHeight, yPos)
		    System.DebugLog("Footer rendered. yPos: " + yPos.ToString)
		    
		    // Save
		    System.DebugLog("Saving PDF...")
		    pdf.Save(saveFile)
		    System.DebugLog("=== PDF SAVED SUCCESSFULLY ===")
		    Return True
		    
		  Catch e As RuntimeException
		    System.DebugLog("ERROR: " + e.Message)
		    If e.Stack <> Nil Then
		      System.DebugLog("Stack: " + String.FromArray(e.Stack, EndOfLine))
		    End If
		    Return False
		  End Try
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetMethodMetrics(methods() As CodeElement, analyzer As ProjectAnalyzer) As MethodMetrics()
		  // Calculate complexity metrics for all methods
		  
		  Var metrics() As MethodMetrics
		  
		  For Each method As CodeElement In methods
		    Var m As New MethodMetrics
		    m.MethodName = method.FullPath
		    m.Element = method
		    
		    // Lines of Code
		    If method.Code.Trim <> "" Then
		      Var lines() As String = method.Code.Split(EndOfLine)
		      Var nonEmptyLines As Integer = 0
		      For Each line As String In lines
		        If line.Trim <> "" Then
		          nonEmptyLines = nonEmptyLines + 1
		        End If
		      Next
		      m.LinesOfCode = nonEmptyLines
		    Else
		      m.LinesOfCode = 0
		    End If
		    
		    // Cyclomatic Complexity
		    If method.Code.Trim <> "" Then
		      m.Complexity = CalculateCyclomaticComplexity(method.Code)
		    Else
		      m.Complexity = 1
		    End If
		    
		    // Call Depth
		    m.CallDepth = method.CallsTo.Count
		    
		    // Call Chain Depth
		    Var visited As New Dictionary
		    m.CallChainDepth = CalculateCallChainDepth(method, visited, 0)
		    
		    // Parameter Complexity
		    If method.Code.Trim <> "" Then
		      Var paramInfo As Dictionary = analyzer.ParseMethodParameters(method.Code)
		      m.ParameterCount = paramInfo.Value("parameterCount")
		      m.OptionalParameterCount = paramInfo.Value("optionalCount")
		      m.HasTooManyParameters = (m.ParameterCount > 5)
		    Else
		      m.ParameterCount = 0
		      m.OptionalParameterCount = 0
		      m.HasTooManyParameters = False
		    End If
		    
		    metrics.Add(m)
		  Next
		  
		  Return metrics
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderCodeSmells(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  //Private Function RenderCodeSmells(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  ' Render code smells section in PDF - SHOW ALL DETAILS
		  
		  Var smells() As CodeSmell = analyzer.DetectCodeSmells()
		  
		  yPos = yPos + 20
		  
		  ' Section header line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, pageWidth - margin, yPos)
		  yPos = yPos + 20
		  
		  ' Title
		  g.FontSize = 16
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  g.DrawText("CODE SMELL DETECTION", margin, yPos)
		  yPos = yPos + 25
		  
		  If smells.Count = 0 Then
		    g.FontSize = 11
		    g.Bold = False
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText(" No code smells detected - excellent code quality!", margin + 20, yPos)
		    yPos = yPos + lineHeight + 20
		    Return yPos
		  End If
		  
		  ' Summary stats
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  g.DrawText("Total Code Smells: " + smells.Count.ToString, margin, yPos)
		  yPos = yPos + lineHeight + 10
		  
		  ' Count by severity
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
		  
		  ' Severity breakdown
		  g.FontSize = 11
		  
		  If critical > 0 Then
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.DrawText("  CRITICAL: " + critical.ToString + " (fix immediately)", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  If high > 0 Then
		    g.DrawingColor = Color.RGB(220, 100, 0)
		    g.DrawText("  HIGH: " + high.ToString + " (fix soon)", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  If medium > 0 Then
		    g.DrawingColor = Color.RGB(180, 180, 0)
		    g.DrawText("  MEDIUM: " + medium.ToString + " (address when possible)", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  If low > 0 Then
		    g.DrawingColor = Color.RGB(100, 150, 100)
		    g.DrawText("  LOW: " + low.ToString + " (nice to fix)", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  yPos = yPos + 15
		  
		  ' Count by type
		  Var typeCount As New Dictionary
		  For Each smell As CodeSmell In smells
		    If typeCount.HasKey(smell.SmellType) Then
		      typeCount.Value(smell.SmellType) = CType(typeCount.Value(smell.SmellType), Integer) + 1
		    Else
		      typeCount.Value(smell.SmellType) = 1
		    End If
		  Next
		  
		  g.FontSize = 11
		  g.DrawingColor = Color.RGB(70, 70, 70)
		  g.Bold = True
		  g.DrawText("By Type:", margin, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  For Each key As Variant In typeCount.Keys
		    Var typeName As String = key.StringValue
		    Var count As Integer = CType(typeCount.Value(key), Integer)
		    g.DrawText("  • " + typeName + ": " + count.ToString, margin + 20, yPos)
		    yPos = yPos + lineHeight + 2
		  Next
		  
		  yPos = yPos + 20
		  
		  ' ========================================
		  ' CRITICAL PRIORITY - ALL WITH DETAILS
		  ' ========================================
		  If critical > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.DrawText(" CRITICAL PRIORITY (" + critical.ToString + " issues)", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "CRITICAL" Then
		        yPos = RenderSingleCodeSmell(g, smell, pageWidth, margin, lineHeight, yPos)
		      End If
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  ' ========================================
		  ' HIGH PRIORITY - ALL WITH DETAILS
		  ' ========================================
		  If high > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(220, 100, 0)
		    g.DrawText("HIGH PRIORITY (" + high.ToString + " issues)", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "HIGH" Then
		        yPos = RenderSingleCodeSmell(g, smell, pageWidth, margin, lineHeight, yPos)
		      End If
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  ' ========================================
		  ' MEDIUM PRIORITY - ALL WITH DETAILS
		  ' ========================================
		  If medium > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(180, 180, 0)
		    g.DrawText("MEDIUM PRIORITY (" + medium.ToString + " issues)", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "MEDIUM" Then
		        yPos = RenderSingleCodeSmell(g, smell, pageWidth, margin, lineHeight, yPos)
		      End If
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  ' ========================================
		  ' LOW PRIORITY - ALL WITH DETAILS
		  ' ========================================
		  If low > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(100, 150, 100)
		    g.DrawText("LOW PRIORITY (" + low.ToString + " issues)", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    For Each smell As CodeSmell In smells
		      If smell.Severity = "LOW" Then
		        yPos = RenderSingleCodeSmell(g, smell, pageWidth, margin, lineHeight, yPos)
		      End If
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderComplexityMetrics(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render complexity metrics section
		  
		  Var allMethods() As CodeElement = analyzer.GetMethodElements
		  Var methodMetrics() As MethodMetrics = GetMethodMetrics(allMethods, analyzer)
		  
		  yPos = yPos + 20
		  
		  // Section header
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, 612 - margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.RGB(150, 0, 150)
		  g.DrawText("COMPLEXITY METRICS", margin, yPos)
		  yPos = yPos + 25
		  
		  // Calculate aggregate stats
		  Var stats As Dictionary = CalculateAggregateStats(methodMetrics)
		  
		  // Display stats - convert integers to strings properly
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  Var totalLOC As Integer = stats.Value("totalLOC")
		  g.DrawText("Total Lines of Code: " + totalLOC.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  Var avgLOC As Integer = stats.Value("avgLOC")
		  g.DrawText("Average Method Size: " + avgLOC.ToString + " lines", margin, yPos)
		  yPos = yPos + lineHeight
		  
		  Var avgComplexity As Integer = stats.Value("avgComplexity")
		  g.DrawText("Average Cyclomatic Complexity: " + avgComplexity.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  Var maxComplexity As Integer = stats.Value("maxComplexity")
		  g.DrawText("Maximum Complexity: " + maxComplexity.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  Var maxLOC As Integer = stats.Value("maxLOC")
		  g.DrawText("Largest Method: " + maxLOC.ToString + " lines", margin, yPos)
		  yPos = yPos + lineHeight
		  
		  Var maxChainDepth As Integer = stats.Value("maxChainDepth")
		  g.DrawText("Deepest Call Chain: " + maxChainDepth.ToString + " levels", margin, yPos)
		  yPos = yPos + lineHeight
		  
		  '// Parameter complexity stats
		  'yPos = yPos + 10
		  'g.FontSize = 11
		  'g.Bold = True
		  'g.DrawingColor = Color.RGB(100, 100, 150)
		  'g.DrawText("Parameter Complexity:", margin, yPos)
		  'yPos = yPos + lineHeight + 3
		  '
		  'g.FontSize = 12
		  'g.Bold = False
		  'g.DrawingColor = Color.Black
		  '
		  'Var avgParameters As Integer = stats.Value("avgParameters")
		  'g.DrawText("Average Parameters per Method: " + avgParameters.ToString, margin, yPos)
		  'yPos = yPos + lineHeight
		  
		  'Var maxParameterCount As Integer = stats.Value("maxParameterCount")
		  'g.DrawText("Maximum Parameters: " + maxParameterCount.ToString, margin, yPos)
		  'yPos = yPos + lineHeight
		  
		  'Var methodsWithTooManyParams As Integer = stats.Value("methodsWithTooManyParams")
		  'If methodsWithTooManyParams > 0 Then
		  'g.DrawingColor = Color.RGB(200, 0, 0)
		  'g.DrawText("⚠ Methods with >5 Parameters: " + methodsWithTooManyParams.ToString, margin, yPos)
		  'Else
		  'g.DrawingColor = Color.RGB(0, 150, 0)
		  'g.DrawText("No methods exceed 5 parameters", margin, yPos)
		  'End If
		  'yPos = yPos + lineHeight
		  '
		  'Var totalOptionalParams As Integer = stats.Value("totalOptionalParams")
		  'g.DrawingColor = Color.Black
		  'g.DrawText("Total Optional Parameters: " + totalOptionalParams.ToString, margin, yPos)
		  yPos = yPos + 25
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderErrorHandlingAnalysis(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderErrorHandlingAnalysis(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  ' Render complete error handling analysis with ALL methods
		  
		  yPos = yPos + 20
		  
		  ' Section header line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, margin + 400, yPos)
		  yPos = yPos + 20
		  
		  ' Title
		  g.FontSize = 16
		  g.Bold = True
		  g.DrawingColor = Color.RGB(200, 50, 50)
		  g.DrawText("[!] ERROR HANDLING ANALYSIS", margin, yPos)
		  yPos = yPos + lineHeight + 15
		  
		  ' Get error handling stats
		  Var stats As Dictionary = analyzer.GetErrorHandlingStats()
		  
		  Var totalMethods As Integer = 0
		  Var methodsWithTryCatch As Integer = 0
		  Var coveragePercent As Integer = 0
		  Var riskyMethods() As CodeElement
		  Var highRiskCount As Integer = 0
		  Var mediumRiskCount As Integer = 0
		  
		  If stats.HasKey("totalMethods") Then
		    totalMethods = stats.Value("totalMethods")
		  End If
		  
		  If stats.HasKey("methodsWithTryCatch") Then
		    methodsWithTryCatch = stats.Value("methodsWithTryCatch")
		  End If
		  
		  If stats.HasKey("coveragePercent") Then
		    coveragePercent = stats.Value("coveragePercent")
		  End If
		  
		  If stats.HasKey("riskyMethods") Then
		    riskyMethods = stats.Value("riskyMethods")
		  End If
		  
		  If stats.HasKey("highRiskCount") Then
		    highRiskCount = stats.Value("highRiskCount")
		  End If
		  
		  If stats.HasKey("mediumRiskCount") Then
		    mediumRiskCount = stats.Value("mediumRiskCount")
		  End If
		  
		  ' Separate methods by highest risk level
		  Var highRiskMethods() As CodeElement
		  Var mediumRiskMethods() As CodeElement
		  
		  For Each method As CodeElement In riskyMethods
		    Var hasHighRisk As Boolean = False
		    
		    For Each pattern As ErrorPattern In method.RiskyPatterns
		      If pattern.RiskLevel = "HIGH" Then
		        hasHighRisk = True
		        Exit For
		      End If
		    Next
		    
		    If hasHighRisk Then
		      highRiskMethods.Add(method)
		    Else
		      mediumRiskMethods.Add(method)
		    End If
		  Next
		  
		  ' Summary stats
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  g.DrawText("Total Methods Analyzed: " + totalMethods.ToString, margin, yPos)
		  yPos = yPos + lineHeight + 3
		  
		  g.DrawText("Methods with Try/Catch: " + methodsWithTryCatch.ToString, margin, yPos)
		  yPos = yPos + lineHeight + 3
		  
		  g.DrawText("Error Handling Coverage: " + coveragePercent.ToString + "%", margin, yPos)
		  yPos = yPos + lineHeight + 3
		  
		  g.DrawText("Methods Without Error Handling: " + riskyMethods.Count.ToString, margin, yPos)
		  yPos = yPos + lineHeight + 10
		  
		  ' Risk breakdown
		  g.FontSize = 11
		  
		  If highRiskMethods.Count > 0 Then
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.DrawText("  • HIGH Risk Issues: " + highRiskMethods.Count.ToString, margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  If mediumRiskMethods.Count > 0 Then
		    g.DrawingColor = Color.RGB(220, 140, 0)
		    g.DrawText("  • MEDIUM Risk Issues: " + mediumRiskMethods.Count.ToString, margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  yPos = yPos + 20
		  
		  ' ========================================
		  ' HIGH RISK METHODS - ALL WITH DETAILS
		  ' ========================================
		  If highRiskMethods.Count > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.DrawText("[!!] HIGH RISK METHODS - All " + highRiskMethods.Count.ToString + " Items", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    g.FontSize = 10
		    g.Bold = False
		    
		    For Each method As CodeElement In highRiskMethods
		      ' Method name in red
		      g.DrawingColor = Color.RGB(150, 0, 0)
		      g.Bold = True
		      g.DrawText(" [!] " + method.FullPath, margin + 20, yPos)
		      yPos = yPos + lineHeight + 2
		      
		      ' Show all risky patterns for this method
		      g.FontSize = 9
		      g.Bold = False
		      g.DrawingColor = Color.RGB(100, 100, 100)
		      
		      For Each pattern As ErrorPattern In method.RiskyPatterns
		        If pattern.RiskLevel = "HIGH" Then
		          g.DrawText("   " + pattern.Description, margin + 25, yPos)
		          yPos = yPos + lineHeight + 2
		        End If
		      Next
		      
		      yPos = yPos + 6
		      g.FontSize = 10
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  ' ========================================
		  ' MEDIUM RISK METHODS - ALL WITH DETAILS
		  ' ========================================
		  If mediumRiskMethods.Count > 0 Then
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.RGB(220, 140, 0)
		    g.DrawText("[!] MEDIUM RISK METHODS - All " + mediumRiskMethods.Count.ToString + " Items", margin, yPos)
		    yPos = yPos + lineHeight + 10
		    
		    g.FontSize = 10
		    g.Bold = False
		    
		    For Each method As CodeElement In mediumRiskMethods
		      ' Method name in orange
		      g.DrawingColor = Color.RGB(180, 100, 0)
		      g.Bold = True
		      g.DrawText(" [!] " + method.FullPath, margin + 20, yPos)
		      yPos = yPos + lineHeight + 2
		      
		      ' Show all risky patterns for this method
		      g.FontSize = 9
		      g.Bold = False
		      g.DrawingColor = Color.RGB(100, 100, 100)
		      
		      For Each pattern As ErrorPattern In method.RiskyPatterns
		        g.DrawText("   " + pattern.Description, margin + 25, yPos)
		        yPos = yPos + lineHeight + 2
		      Next
		      
		      yPos = yPos + 6
		      g.FontSize = 10
		    Next
		    
		    yPos = yPos + 10
		  End If
		  
		  ' Success message if no issues
		  If riskyMethods.Count = 0 Then
		    g.FontSize = 12
		    g.Bold = False
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText("[OK] All methods have appropriate error handling!", margin + 20, yPos)
		    yPos = yPos + lineHeight + 20
		  End If
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderFooter(g As Graphics, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render report footer
		  
		  yPos = yPos + 30
		  
		  // Decorative line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, pageWidth - margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 11
		  g.Bold = False
		  g.DrawingColor = Color.RGB(0, 100, 0)
		  DrawCenteredText(g, "Scan completed successfully!", pageWidth / 2, yPos, pageWidth - (margin * 2))
		  yPos = yPos + lineHeight
		  
		  // Timestamp
		  yPos = yPos + 20
		  g.FontSize = 9
		  g.DrawingColor = Color.RGB(128, 128, 128)
		  Var dt As DateTime = DateTime.Now
		  Var timestamp As String = dt.ToString
		  DrawCenteredText(g, "Generated: " + timestamp, pageWidth / 2, yPos, pageWidth - (margin * 2))
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderHeader(g As Graphics, pageWidth As Integer, margin As Double, yPos As Double) As Double
		  // Render the report header
		  
		  g.FontSize = 18
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  DrawCenteredText(g, "CODE ANALYSIS REPORT", pageWidth / 2, yPos, pageWidth - (margin * 2))
		  yPos = yPos + 30
		  
		  // Decorative line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, pageWidth - margin, yPos)
		  yPos = yPos + 25
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderParameterComplexityDetails(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderParameterComplexity(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render parameter complexity analysis
		  
		  Var currentY As Double = yPos
		  
		  // Section header
		  g.DrawingColor = Color.RGB(70, 130, 180)
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawText("PARAMETER COMPLEXITY", margin, currentY)
		  g.Bold = False
		  currentY = currentY + lineHeight * 1.5
		  
		  // Get all methods
		  Var methods() As CodeElement = analyzer.GetMethodElements
		  
		  If methods.Count = 0 Then
		    g.FontSize = 11
		    g.DrawingColor = Color.RGB(100, 100, 100)
		    g.DrawText("No methods found for analysis", margin + 20, currentY)
		    Return currentY + lineHeight
		  End If
		  
		  // Calculate statistics using STORED parameter data
		  Var totalParams As Integer = 0
		  Var totalOptionalParams As Integer = 0
		  Var maxParams As Integer = 0
		  Var methodsWithManyParams() As CodeElement
		  
		  For Each method As CodeElement In methods
		    totalParams = totalParams + method.ParameterCount
		    totalOptionalParams = totalOptionalParams + method.OptionalParameterCount
		    
		    If method.ParameterCount > maxParams Then
		      maxParams = method.ParameterCount
		    End If
		    
		    If method.ParameterCount > 5 Then
		      methodsWithManyParams.Add(method)
		    End If
		  Next
		  
		  // Calculate average
		  Var avgParams As Double = 0
		  If methods.Count > 0 Then
		    avgParams = totalParams / methods.Count
		  End If
		  
		  // Render statistics
		  g.FontSize = 11
		  g.DrawingColor = Color.Black
		  
		  g.DrawText("Average Parameters per Method: " + avgParams.ToString("0.0"), margin + 20, currentY)
		  currentY = currentY + lineHeight
		  
		  g.DrawText("Maximum Parameters: " + maxParams.ToString, margin + 20, currentY)
		  currentY = currentY + lineHeight
		  
		  Var excessMessage As String
		  If methodsWithManyParams.Count = 0 Then
		    excessMessage = "✓ No methods exceed 5 parameters"
		    g.DrawingColor = Color.RGB(0, 128, 0)
		  Else
		    excessMessage = "⚠ " + methodsWithManyParams.Count.ToString + " method(s) exceed 5 parameters"
		    g.DrawingColor = Color.RGB(200, 100, 0)
		  End If
		  g.DrawText(excessMessage, margin + 20, currentY)
		  currentY = currentY + lineHeight
		  
		  g.DrawingColor = Color.Black
		  g.DrawText("Total Optional Parameters: " + totalOptionalParams.ToString, margin + 20, currentY)
		  currentY = currentY + lineHeight * 1.5
		  
		  // List methods with many parameters
		  If methodsWithManyParams.Count > 0 Then
		    g.DrawingColor = Color.RGB(200, 100, 0)
		    g.Bold = True
		    g.DrawText("Methods with > 5 Parameters:", margin + 20, currentY)
		    g.Bold = False
		    currentY = currentY + lineHeight
		    
		    g.DrawingColor = Color.Black
		    For Each method As CodeElement In methodsWithManyParams
		      Var methodInfo As String = "  • " + method.FullPath + " (" + method.ParameterCount.ToString + " parameters)"
		      g.DrawText(methodInfo, margin + 40, currentY)
		      currentY = currentY + lineHeight
		      
		      // Show the actual parameters
		      If method.Parameters <> "" Then
		        g.FontSize = 10
		        g.DrawingColor = Color.RGB(100, 100, 100)
		        
		        // Split long parameter lists across multiple lines if needed
		        Var paramText As String = "    Parameters: " + method.Parameters
		        If paramText.Length > 100 Then
		          // Wrap long parameter lists
		          Var words() As String = paramText.Split(",")
		          Var currentLine As String = words(0)
		          
		          Var lineLength As String
		          
		          For i As Integer = 1 To words.Count - 1
		            lineLength = currentLine + "," + words(i)
		            
		            If lineLength.Length > 100 Then
		              g.DrawText(currentLine + ",", margin + 40, currentY)
		              currentY = currentY + lineHeight * 0.9
		              currentLine = "      " + words(i).Trim
		            Else
		              currentLine = currentLine + "," + words(i)
		            End If
		          Next
		          
		          g.DrawText(currentLine, margin + 40, currentY)
		          currentY = currentY + lineHeight * 0.9
		        Else
		          g.DrawText(paramText, margin + 40, currentY)
		          currentY = currentY + lineHeight * 0.9
		        End If
		        
		        g.FontSize = 11
		        g.DrawingColor = Color.Black
		      End If
		      
		      currentY = currentY + lineHeight * 0.3
		    Next
		  End If
		  
		  Return currentY + lineHeight
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderPrioritySection(g As Graphics, sectionTitle As String, suggestions() As RefactoringSuggestion, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double, headerColor As Color) As Double
		  // Private Function RenderPrioritySection(g As Graphics, sectionTitle As String, suggestions() As RefactoringSuggestion, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double, headerColor As Color) As Double
		  // Section header
		  g.FontSize = 16
		  g.Bold = True
		  g.DrawingColor = headerColor
		  g.DrawText(sectionTitle, margin, yPos)
		  yPos = yPos + 25
		  
		  // Separator line
		  g.FillRectangle(margin, yPos, pageWidth - (margin * 2), 1)
		  yPos = yPos + 20
		  
		  // Render each suggestion
		  For Each suggestion As RefactoringSuggestion In suggestions
		    yPos = RenderSingleSuggestion(g, suggestion, pageWidth, margin, lineHeight, yPos)
		  Next
		  
		  yPos = yPos + 20
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderQualityScore(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderQualityScore(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  ' Render quality score section in PDF
		  
		  Var score As QualityScore = analyzer.CalculateQualityScore()
		  
		  yPos = yPos + 20
		  
		  ' Section header line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, pageWidth - margin, yPos)
		  yPos = yPos + 20
		  
		  ' Title
		  g.FontSize = 16
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  g.DrawText("CODE QUALITY SCORE", margin, yPos)
		  yPos = yPos + 50
		  
		  ' Overall Score - Large and prominent
		  g.FontSize = 48
		  g.Bold = True
		  g.DrawingColor = score.GetScoreColor()
		  
		  Var scoreText As String = score.OverallScore.ToString("0.0")
		  Var scoreWidth As Double = g.TextWidth(scoreText)
		  Var centerX As Double = pageWidth / 2
		  g.DrawText(scoreText, centerX - (scoreWidth / 2), yPos)
		  yPos = yPos + 50
		  
		  ' Grade
		  g.FontSize = 24
		  Var gradeText As String = "Grade: " + score.Grade
		  Var gradeWidth As Double = g.TextWidth(gradeText)
		  g.DrawText(gradeText, centerX - (gradeWidth / 2), yPos)
		  yPos = yPos + 35
		  
		  ' Status message
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  Var statusMsg As String
		  If score.OverallScore >= 80 Then
		    statusMsg = "EXCELLENT - Well-maintained codebase"
		  ElseIf score.OverallScore >= 60 Then
		    statusMsg = "GOOD - Minor improvements recommended"
		  ElseIf score.OverallScore >= 40 Then
		    statusMsg = "NEEDS WORK - Several issues to address"
		  Else
		    statusMsg = "CRITICAL - Immediate attention required"
		  End If
		  
		  Var statusWidth As Double = g.TextWidth(statusMsg)
		  g.DrawText(statusMsg, centerX - (statusWidth / 2), yPos)
		  yPos = yPos + 30
		  
		  ' Component breakdown
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.RGB(70, 70, 70)
		  g.DrawText("Component Scores (Weighted)", margin, yPos)
		  yPos = yPos + 25
		  
		  g.FontSize = 11
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  ' Error Handling
		  yPos = RenderScoreBar(g, "Error Handling", score.ErrorHandlingScore, 30, _
		  score.ErrorHandlingCoverage.ToString("0.0") + "% coverage", _
		  pageWidth, margin, lineHeight, yPos)
		  yPos = yPos + 5
		  
		  ' Complexity
		  yPos = RenderScoreBar(g, "Complexity", score.ComplexityScore, 25, _
		  "Avg: " + score.AverageComplexity.ToString("0.1"), _
		  pageWidth, margin, lineHeight, yPos)
		  yPos = yPos + 5
		  
		  ' Code Reuse
		  yPos = RenderScoreBar(g, "Code Reuse", score.CodeReuseScore, 20, _
		  score.UnusedPercentage.ToString("0.1") + "% unused", _
		  pageWidth, margin, lineHeight, yPos)
		  yPos = yPos + 5
		  
		  ' Parameters
		  yPos = RenderScoreBar(g, "Parameters", score.ParameterScore, 15, _
		  "Avg: " + score.AverageParameters.ToString("0.1"), _
		  pageWidth, margin, lineHeight, yPos)
		  yPos = yPos + 5
		  
		  ' Documentation
		  yPos = RenderScoreBar(g, "Documentation", score.DocumentationScore, 10, _
		  score.DocumentationCoverage.ToString("0.0") + "% coverage", _
		  pageWidth, margin, lineHeight, yPos)
		  yPos = yPos + 25
		  
		  ' Recommendations
		  g.FontSize = 12
		  g.Bold = True
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawText("Improvement Recommendations:", margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 10
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  Var hasRecommendations As Boolean = False
		  
		  If score.ErrorHandlingScore < 70 Then
		    g.DrawText("• Add try/catch blocks to risky operations", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		    hasRecommendations = True
		  End If
		  
		  If score.ComplexityScore < 70 Then
		    g.DrawText("• Refactor complex methods to improve maintainability", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		    hasRecommendations = True
		  End If
		  
		  If score.CodeReuseScore < 70 Then
		    g.DrawText("• Remove unused code to reduce maintenance burden", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		    hasRecommendations = True
		  End If
		  
		  If score.ParameterScore < 70 Then
		    g.DrawText("• Reduce parameter counts using parameter objects", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		    hasRecommendations = True
		  End If
		  
		  If score.DocumentationScore < 70 Then
		    g.DrawText("• Add comments to explain complex logic", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		    hasRecommendations = True
		  End If
		  
		  If Not hasRecommendations Then
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText("Excellent! No major improvements needed.", margin + 20, yPos)
		    yPos = yPos + lineHeight + 3
		  End If
		  
		  yPos = yPos + 15
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderRefactoringFooter(g As Graphics, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderRefactoringFooter(g As Graphics, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Separator line
		  g.DrawingColor = &c1E3A8A
		  g.FillRectangle(margin, yPos, pageWidth - (margin * 2), 1)
		  yPos = yPos + 15
		  
		  // Footer text
		  g.FontSize = 9
		  g.Bold = False
		  g.DrawingColor = &c666666
		  
		  Var now As DateTime = DateTime.Now
		  Var timestamp As String = now.ToString("d MMM yyyy 'at' h:mm:ss tt", Locale.Current)
		  
		  g.DrawText("Generated: " + timestamp, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  DrawCenteredText(g, "Review and prioritize these refactoring suggestions to improve code quality", pageWidth / 2, yPos, pageWidth - (margin * 2))
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderRefactoringHeader(g As Graphics, pageWidth As Integer, margin As Double, yPos As Double) As Double
		  // Private Function RenderRefactoringHeader(g As Graphics, pageWidth As Integer, margin As Double, yPos As Double) As Double
		  // Title with accent color
		  g.FontSize = 24
		  g.Bold = True
		  g.DrawingColor = &c1E3A8A  // Deep red for emphasis
		  DrawCenteredText(g, "REFACTORING SUGGESTIONS", pageWidth / 2, yPos, pageWidth - (margin * 2))
		  yPos = yPos + 30
		  
		  // Subtitle
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = &c666666
		  DrawCenteredText(g, "Actionable recommendations to improve code quality", pageWidth / 2, yPos, pageWidth - (margin * 2))
		  yPos = yPos + 40
		  
		  // Separator line
		  g.DrawingColor = &c1E3A8A
		  g.FillRectangle(margin, yPos, pageWidth - (margin * 2), 2)
		  yPos = yPos + 30
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderRefactoringSuggestions(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  
		  
		  // Private Function RenderRefactoringSuggestions(g As Graphics, analyzer As ProjectAnalyzer, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Collect all suggestions and group by priority
		  Var highPriority() As RefactoringSuggestion
		  Var mediumPriority() As RefactoringSuggestion
		  Var lowPriority() As RefactoringSuggestion
		  
		  Var methods() As CodeElement = analyzer.GetMethodElements()
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
		  
		  // DEBUG: Check counts BEFORE deduplication
		  MessageBox("Before dedup - High: " + highPriority.Count.ToString + _
		  " Medium: " + mediumPriority.Count.ToString + _
		  " Low: " + lowPriority.Count.ToString)
		  
		  // DEDUPLICATE EACH PRIORITY GROUP
		  DeduplicateSuggestionsInPlace(highPriority)
		  DeduplicateSuggestionsInPlace(mediumPriority)
		  DeduplicateSuggestionsInPlace(lowPriority)
		  
		  // DEBUG: Check counts AFTER deduplication
		  MessageBox("After dedup - High: " + highPriority.Count.ToString + _
		  " Medium: " + mediumPriority.Count.ToString + _
		  " Low: " + lowPriority.Count.ToString)
		  
		  // Render high priority first
		  If highPriority.Count > 0 Then
		    yPos = RenderPrioritySection(g, "HIGH PRIORITY - Fix These First!", highPriority, pageWidth, margin, lineHeight, yPos, &cDC143C)
		  End If
		  
		  // Then medium
		  If mediumPriority.Count > 0 Then
		    yPos = RenderPrioritySection(g, "MEDIUM PRIORITY", mediumPriority, pageWidth, margin, lineHeight, yPos, &cFF8C00)
		  End If
		  
		  // Then low (if any)
		  If lowPriority.Count > 0 Then
		    yPos = RenderPrioritySection(g, "LOW PRIORITY", lowPriority, pageWidth, margin, lineHeight, yPos, &c4169E1)
		  End If
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderRefactoringSummary(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderRefactoringSummary(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Collect and categorize suggestions
		  Var highPriority() As RefactoringSuggestion
		  Var mediumPriority() As RefactoringSuggestion
		  Var lowPriority() As RefactoringSuggestion
		  
		  Var methods() As CodeElement = analyzer.GetMethodElements()
		  
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
		  
		  // DEDUPLICATE BEFORE COUNTING - This is the fix!
		  DeduplicateSuggestionsInPlace(highPriority)
		  DeduplicateSuggestionsInPlace(mediumPriority)
		  DeduplicateSuggestionsInPlace(lowPriority)
		  
		  // Calculate total from deduplicated arrays
		  Var totalIssues As Integer = highPriority.Count + mediumPriority.Count + lowPriority.Count
		  
		  // Summary box
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  g.DrawText("Summary", margin, yPos)
		  yPos = yPos + lineHeight + 10
		  
		  g.FontSize = 12
		  g.Bold = False
		  
		  g.DrawText("Total Issues Found: " + totalIssues.ToString, margin + 20, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  // High priority in red
		  g.DrawingColor = &cDC143C  // Crimson red
		  g.DrawText("  [HIGH] High Priority: " + highPriority.Count.ToString + " (Fix First!)", margin + 20, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  // Medium priority in orange
		  g.DrawingColor = &cFF8C00  // Dark orange
		  g.DrawText("  [Medium] Medium Priority: " + mediumPriority.Count.ToString, margin + 20, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  // Low priority in blue (if any)
		  If lowPriority.Count > 0 Then
		    g.DrawingColor = &c4169E1  // Royal blue
		    g.DrawText("  [Low] Low Priority: " + lowPriority.Count.ToString, margin + 20, yPos)
		    yPos = yPos + lineHeight + 5
		  End If
		  
		  yPos = yPos + 20
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderRelationships(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render relationship analysis section
		  
		  Var allMethods() As CodeElement = analyzer.GetMethodElements
		  
		  yPos = yPos + 20
		  
		  // Section header
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, 612 - margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  g.DrawText("RELATIONSHIP ANALYSIS", margin, yPos)
		  yPos = yPos + 25
		  
		  // Count methods with calls
		  Var methodsWithCalls As Integer = 0
		  Var totalCalls As Integer = 0
		  
		  For Each element As CodeElement In allMethods
		    If element.CallsTo.Count > 0 Then
		      methodsWithCalls = methodsWithCalls + 1
		      totalCalls = totalCalls + element.CallsTo.Count
		    End If
		  Next
		  
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawText("Methods with outgoing calls: " + methodsWithCalls.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  g.DrawText("Total method calls detected: " + totalCalls.ToString, margin, yPos)
		  yPos = yPos + 20
		  
		  // Sample relationships
		  If totalCalls > 0 Then
		    yPos = RenderSampleRelationships(g, allMethods, margin, lineHeight, yPos)
		  End If
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderSampleRelationships(g As Graphics, allMethods() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render sample method call relationships
		  
		  g.FontSize = 11
		  g.Bold = True
		  Var maxSamples As Integer = 15
		  g.DrawText("Sample relationships (first " + maxSamples.ToString + "):", margin, yPos)
		  yPos = yPos + lineHeight + 3
		  
		  g.FontSize = 10
		  g.Bold = False
		  g.DrawingColor = Color.RGB(50, 50, 150)
		  
		  Var sampleCount As Integer = 0
		  For Each element As CodeElement In allMethods
		    If element.CallsTo.Count > 0 Then
		      For Each calledElement As CodeElement In element.CallsTo
		        If sampleCount >= maxSamples Then
		          Exit For element
		        End If
		        
		        g.DrawText("  " + element.FullPath + " -> " + calledElement.FullPath, margin + 10, yPos)
		        yPos = yPos + lineHeight
		        sampleCount = sampleCount + 1
		      Next
		    End If
		  Next
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderScoreBar(g As Graphics, label As String, score As Double, weight As Integer, detail As String, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  //Private Function RenderScoreBar(g As Graphics, label As String, score As Double, weight As Integer, detail As String, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  ' Render a score bar with label and percentage
		  
		  Var barWidth As Double = 280  // Reduced from 300 to give more room
		  Var barHeight As Double = 20
		  Var barX As Double = margin + 150
		  
		  ' Label
		  g.FontSize = 11
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  g.DrawText(label + " (" + weight.ToString + "%):", margin, yPos + 15)
		  
		  ' Score bar background
		  g.DrawingColor = Color.RGB(220, 220, 220)
		  g.FillRectangle(barX, yPos, barWidth, barHeight)
		  
		  ' Score bar fill (colored based on score)
		  Var fillWidth As Double = (score / 100.0) * barWidth
		  
		  If score >= 80 Then
		    g.DrawingColor = Color.RGB(0, 180, 0)  // Green
		  ElseIf score >= 60 Then
		    g.DrawingColor = Color.RGB(100, 180, 0)  // Yellow-green
		  ElseIf score >= 40 Then
		    g.DrawingColor = Color.RGB(255, 140, 0)  // Orange
		  Else
		    g.DrawingColor = Color.RGB(220, 50, 50)  // Red
		  End If
		  
		  g.FillRectangle(barX, yPos, fillWidth, barHeight)
		  
		  ' Border
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawRectangle(barX, yPos, barWidth, barHeight)
		  
		  ' Score text (right aligned to bar)
		  g.FontSize = 10
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  Var scoreText As String = score.ToString("0.0")
		  g.DrawText(scoreText, barX + barWidth + 10, yPos + 15)
		  
		  ' Detail text on next line (moved down)
		  g.FontSize = 9
		  g.Bold = False
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawText(detail, barX + 10, yPos + barHeight + 12)  // ← MOVED TO NEW LINE
		  
		  Return yPos + barHeight + 18  // ← Adjusted spacing to account for detail line
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderSingleCodeSmell(g As Graphics, smell As CodeSmell, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderSingleCodeSmell(g As Graphics, smell As CodeSmell, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  
		  ' Render a single code smell with formatting
		  
		  ' Element path with severity indicator
		  g.FontSize = 11
		  g.Bold = True
		  g.DrawingColor = smell.GetSeverityColor()
		  g.DrawText("• " + smell.SmellType, margin + 20, yPos)
		  yPos = yPos + lineHeight + 2
		  
		  ' Location
		  g.FontSize = 10
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  g.DrawText("  " + smell.Element.FullPath, margin + 25, yPos)
		  yPos = yPos + lineHeight + 2
		  
		  ' Description
		  g.FontSize = 9
		  g.DrawingColor = Color.RGB(80, 80, 80)
		  Var wrappedDesc() As String = WrapText(g, smell.Description, pageWidth - margin - 60)
		  For Each line As String In wrappedDesc
		    g.DrawText("  " + line, margin + 25, yPos)
		    yPos = yPos + lineHeight
		  Next
		  
		  ' Details
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  Var wrappedDetails() As String = WrapText(g, smell.Details, pageWidth - margin - 60)
		  For Each line As String In wrappedDetails
		    g.DrawText("  " + line, margin + 25, yPos)
		    yPos = yPos + lineHeight
		  Next
		  
		  ' Recommendation
		  g.DrawingColor = Color.RGB(0, 100, 150)
		  g.DrawText("  --> " + smell.Recommendation, margin + 25, yPos)
		  yPos = yPos + lineHeight + 10
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderSingleSuggestion(g As Graphics, suggestion As RefactoringSuggestion, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderSingleSuggestion(g As Graphics, suggestion As RefactoringSuggestion, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Safety checks
		  If suggestion = Nil Then
		    System.DebugLog("WARNING: Suggestion is Nil")
		    Return yPos
		  End If
		  
		  If suggestion.Element = Nil Then
		    System.DebugLog("WARNING: Suggestion.Element is Nil")
		    Return yPos
		  End If
		  
		  // Method name in bold
		  g.FontSize = 11
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  g.DrawText("--> " + suggestion.Element.FullPath, margin + 10, yPos)
		  yPos = yPos + lineHeight + 3
		  // Issue title
		  g.FontSize = 10
		  g.Bold = True
		  g.DrawingColor = &c444444
		  g.DrawText("   " + suggestion.Title, margin + 10, yPos)
		  yPos = yPos + lineHeight + 2
		  
		  // Description
		  g.FontSize = 9
		  g.Bold = False
		  g.DrawingColor = &c666666
		  g.DrawText("   " + suggestion.Description, margin + 10, yPos)
		  yPos = yPos + lineHeight + 8
		  
		  // Suggestions with checkboxes
		  g.FontSize = 9
		  g.DrawingColor = Color.Black
		  
		  For Each tip As String In suggestion.Suggestions
		    If tip.Trim = "" Then
		      yPos = yPos + 3  // Small space for blank lines
		      Continue
		    End If
		    
		    // Wrap long lines
		    Var maxWidth As Double = pageWidth - margin - 40
		    Var wrappedLines() As String = WrapText(g, tip, maxWidth)
		    
		    For i As Integer = 0 To wrappedLines.LastIndex
		      If i = 0 Then
		        // First line with checkbox
		        g.DrawText("  [  ]  " + wrappedLines(i), margin + 20, yPos)
		      Else
		        // Continuation lines indented
		        g.DrawText("      " + wrappedLines(i), margin + 20, yPos)
		      End If
		      yPos = yPos + lineHeight
		    Next
		  Next
		  
		  yPos = yPos + 15  // Space between suggestions
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderSummary(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Private Function RenderSummary(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render summary statistics section
		  
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.Black
		  
		  Var allElements() As CodeElement = analyzer.GetAllElements()
		  Var totalCount As String = allElements.Count.ToString
		  g.DrawText("Total Elements Found: " + totalCount, margin, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  g.FontSize = 12
		  g.Bold = False
		  
		  Var classElements() As CodeElement = analyzer.GetClassElements()
		  Var classCount As String = classElements.Count.ToString
		  g.DrawText("  - Classes: " + classCount, margin + 20, yPos)
		  yPos = yPos + lineHeight
		  
		  Var moduleElements() As CodeElement = analyzer.GetModuleElements()
		  Var moduleCount As String = moduleElements.Count.ToString
		  g.DrawText("  - Modules: " + moduleCount, margin + 20, yPos)
		  yPos = yPos + lineHeight
		  
		  Var methodElements() As CodeElement = analyzer.GetMethodElements()
		  Var methodCount As String = methodElements.Count.ToString
		  g.DrawText("  - Methods: " + methodCount, margin + 20, yPos)
		  yPos = yPos + 25
		  
		  Return yPos
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderTopComplexMethods(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render top 10 most complex methods
		  
		  Var allMethods() As CodeElement = analyzer.GetMethodElements
		  Var methodMetrics() As MethodMetrics = GetMethodMetrics(allMethods, analyzer)
		  Var sortedMetrics() As MethodMetrics = SortMethodsByComplexity(methodMetrics)
		  
		  g.FontSize = 13
		  g.Bold = True
		  g.DrawingColor = Color.RGB(200, 0, 0)
		  g.DrawText("MOST COMPLEX METHODS (Top 10)", margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 10
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  Var displayCount As Integer = 10
		  If sortedMetrics.Count < displayCount Then
		    displayCount = sortedMetrics.Count
		  End If
		  
		  For i As Integer = 0 To displayCount - 1
		    Var metric As MethodMetrics = sortedMetrics(i)
		    
		    // Method name
		    g.Bold = True
		    Var rankNum As Integer = i + 1
		    g.DrawText(rankNum.ToString + ". " + metric.MethodName, margin + 5, yPos)
		    yPos = yPos + lineHeight
		    
		    // Metrics
		    g.Bold = False
		    g.DrawingColor = Color.RGB(80, 80, 80)
		    
		    Var metricsLine As String = "   LOC: " + metric.LinesOfCode.ToString
		    metricsLine = metricsLine + " | Complexity: " + metric.Complexity.ToString
		    metricsLine = metricsLine + " | Calls: " + metric.CallDepth.ToString
		    metricsLine = metricsLine + " | Chain: " + metric.CallChainDepth.ToString
		    
		    g.DrawText(metricsLine, margin + 10, yPos)
		    yPos = yPos + lineHeight + 3
		    
		    g.DrawingColor = Color.Black
		  Next
		  
		  yPos = yPos + 15
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderUnusedByType(g As Graphics, unusedElements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render unused elements grouped by type
		  
		  // Group by element type
		  Var methodsList() As CodeElement
		  Var classesList() As CodeElement
		  Var modulesList() As CodeElement
		  Var propertiesList() As CodeElement
		  Var variablesList() As CodeElement
		  
		  For Each element As CodeElement In unusedElements
		    Select Case element.ElementType
		    Case "METHOD"
		      methodsList.Add(element)
		    Case "CLASS"
		      classesList.Add(element)
		    Case "MODULE"
		      modulesList.Add(element)
		    Case "PROPERTY"
		      propertiesList.Add(element)
		    Case "VARIABLE"
		      variablesList.Add(element)
		    End Select
		  Next
		  
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  g.FontSize = 11
		  
		  // Render each type
		  If methodsList.Count > 0 Then
		    yPos = RenderUnusedTypeList(g, "METHOD", methodsList, margin, lineHeight, yPos)
		  End If
		  If classesList.Count > 0 Then
		    yPos = RenderUnusedTypeList(g, "CLASS", classesList, margin, lineHeight, yPos)
		  End If
		  If modulesList.Count > 0 Then
		    yPos = RenderUnusedTypeList(g, "MODULE", modulesList, margin, lineHeight, yPos)
		  End If
		  If propertiesList.Count > 0 Then
		    yPos = RenderUnusedTypeList(g, "PROPERTY", propertiesList, margin, lineHeight, yPos)
		  End If
		  If variablesList.Count > 0 Then
		    yPos = RenderUnusedTypeList(g, "VARIABLE", variablesList, margin, lineHeight, yPos)
		  End If
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderUnusedElements(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render unused elements section
		  
		  Var unusedElements() As CodeElement = analyzer.GetUnusedElements()
		  
		  g.FontSize = 14
		  g.Bold = True
		  
		  If unusedElements.Count > 0 Then
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    Var unusedCount As String = unusedElements.Count.ToString
		    g.DrawText("Found " + unusedCount + " UNUSED elements:", margin, yPos)
		  Else
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText("No unused elements found", margin, yPos)
		  End If
		  yPos = yPos + 25
		  
		  If unusedElements.Count > 0 Then
		    yPos = RenderUnusedByType(g, unusedElements, margin, lineHeight, yPos)
		  End If
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function RenderUnusedTypeList(g As Graphics, typeName As String, elements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double
		  // Render a list of unused elements of a specific type
		  
		  g.FontSize = 12
		  g.Bold = True
		  Var count As String = elements.Count.ToString
		  g.DrawText(typeName + " (" + count + "):", margin, yPos)
		  yPos = yPos + lineHeight + 3
		  
		  g.FontSize = 10
		  g.Bold = False
		  For Each element As CodeElement In elements
		    g.DrawText("  • " + element.FullPath + " [" + element.FileName + "]", margin + 10, yPos)
		    yPos = yPos + lineHeight
		  Next
		  yPos = yPos + 10
		  
		  Return yPos
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SortMethodsByComplexity(metrics() As MethodMetrics) As MethodMetrics()
		  // Sort methods by composite complexity score
		  
		  Var sorted() As MethodMetrics
		  For Each m As MethodMetrics In metrics
		    sorted.Add(m)
		  Next
		  
		  // Bubble sort
		  For i As Integer = 0 To sorted.Count - 1
		    For j As Integer = i + 1 To sorted.Count - 1
		      Var score1 As Double = (sorted(i).Complexity * 2.0) + (sorted(i).LinesOfCode / 10.0) + (sorted(i).CallChainDepth * 3.0)
		      Var score2 As Double = (sorted(j).Complexity * 2.0) + (sorted(j).LinesOfCode / 10.0) + (sorted(j).CallChainDepth * 3.0)
		      
		      If score2 > score1 Then
		        Var temp As MethodMetrics = sorted(i)
		        sorted(i) = sorted(j)
		        sorted(j) = temp
		      End If
		    Next
		  Next
		  
		  Return sorted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function WrapText(g As Graphics, text As String, maxWidth As Double) As String()
		  // Private Function WrapText(g As Graphics, text As String, maxWidth As Double) As String()
		  
		  Var lines() As String
		  Var words() As String = Text.Split(" ")
		  Var currentLine As String = ""
		  
		  For Each word As String In words
		    Var testLine As String
		    If currentLine = "" Then
		      testLine = word
		    Else
		      testLine = currentLine + " " + word
		    End If
		    
		    If g.TextWidth(testLine) <= maxWidth Then
		      currentLine = testLine
		    Else
		      If currentLine <> "" Then
		        lines.Add(currentLine)
		      End If
		      currentLine = word
		    End If
		  Next
		  
		  If currentLine <> "" Then
		    lines.Add(currentLine)
		  End If
		  
		  Return lines
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
End Class
#tag EndClass
