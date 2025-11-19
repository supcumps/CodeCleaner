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
		  // Calculate total height needed for the PDF
		  
		  Var unusedElements() As CodeElement = analyzer.GetUnusedElements()
		  Var allMethods() As CodeElement = analyzer.GetMethodElements
		  
		  Var estimatedLines As Integer = 50  // Header + summary
		  estimatedLines = estimatedLines + (unusedElements.Count * 2)
		  estimatedLines = estimatedLines + 40  // Error handling section
		  estimatedLines = estimatedLines + 20  // Relationship section
		  estimatedLines = estimatedLines + 50  // Complexity section (including parameter stats)
		  estimatedLines = estimatedLines + 25  // Parameter complexity details
		  estimatedLines = estimatedLines + 30  // Most complex methods
		  
		  // Count relationship lines
		  Var relationshipCount As Integer = 0
		  For Each element As CodeElement In allMethods
		    If element.CallsTo.Count > 0 Then
		      relationshipCount = relationshipCount + element.CallsTo.Count
		    End If
		  Next
		  If relationshipCount > 15 Then
		    relationshipCount = 15
		  End If
		  estimatedLines = estimatedLines + relationshipCount
		  
		  Return CType(estimatedLines * lineHeight + (margin * 3), Integer)
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
		Private Sub DrawCenteredText(g As Graphics, text As String, centerX As Double, y As Double, maxWidth As Double)
		  // Draw text centered at the given x position
		  Var textWidth As Double = g.TextWidth(text)
		  Var x As Double = centerX - (textWidth / 2)
		  g.DrawText(text, x, y)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GenerateAnalysisReportPDF(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean
		  // REFACTORED VERSION - Much cleaner!
		  
		  Try
		    // Setup
		    Var pageWidth As Integer = 612
		    Var margin As Double = 50
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
		  'g.DrawText("✓ No methods exceed 5 parameters", margin, yPos)
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
		  // Render error handling analysis section
		  
		  Var stats As Dictionary = analyzer.GetErrorHandlingStats()
		  
		  yPos = yPos + 20
		  
		  // Section header
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawLine(margin, yPos, 612 - margin, yPos)
		  yPos = yPos + 20
		  
		  g.FontSize = 14
		  g.Bold = True
		  g.DrawingColor = Color.RGB(200, 0, 0)
		  g.DrawText("⚠ ERROR HANDLING ANALYSIS", margin, yPos)
		  yPos = yPos + 25
		  
		  // Statistics
		  g.FontSize = 12
		  g.Bold = False
		  g.DrawingColor = Color.Black
		  
		  Var totalMethods As Integer = stats.Value("totalMethods")
		  Var methodsWithTryCatch As Integer = stats.Value("methodsWithTryCatch")
		  Var coveragePercent As Integer = stats.Value("coveragePercent")
		  Var highRiskCount As Integer = stats.Value("highRiskCount")
		  Var mediumRiskCount As Integer = stats.Value("mediumRiskCount")
		  Var methodsWithRiskyPatterns As Integer = stats.Value("methodsWithRiskyPatterns")
		  
		  g.DrawText("Total Methods Analyzed: " + totalMethods.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  g.DrawText("Methods with Try/Catch: " + methodsWithTryCatch.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  // Coverage with color coding
		  g.DrawText("Error Handling Coverage: " + coveragePercent.ToString + "%", margin, yPos)
		  If coveragePercent < 30 Then
		    g.DrawingColor = Color.RGB(200, 0, 0)  // Red for poor coverage
		  ElseIf coveragePercent < 70 Then
		    g.DrawingColor = Color.RGB(255, 140, 0)  // Orange for moderate coverage
		  Else
		    g.DrawingColor = Color.RGB(0, 150, 0)  // Green for good coverage
		  End If
		  Var barWidth As Integer = coveragePercent * 2
		  g.FillRectangle(margin + 200, yPos - 10, barWidth, 10)
		  g.DrawingColor = Color.Black
		  yPos = yPos + lineHeight + 5
		  
		  yPos = yPos + 10
		  g.DrawText("Methods Without Error Handling: " + methodsWithRiskyPatterns.ToString, margin, yPos)
		  yPos = yPos + lineHeight
		  
		  g.DrawingColor = Color.RGB(200, 0, 0)
		  g.DrawText("  • HIGH Risk Issues: " + highRiskCount.ToString, margin + 20, yPos)
		  yPos = yPos + lineHeight
		  
		  g.DrawingColor = Color.RGB(255, 140, 0)
		  g.DrawText("  • MEDIUM Risk Issues: " + mediumRiskCount.ToString, margin + 20, yPos)
		  yPos = yPos + lineHeight + 5
		  
		  // High-risk methods list
		  If methodsWithRiskyPatterns > 0 Then
		    yPos = yPos + 10
		    g.FontSize = 13
		    g.Bold = True
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.DrawText("🔴 HIGH RISK METHODS (Need Immediate Attention)", margin, yPos)
		    yPos = yPos + 20
		    
		    g.FontSize = 10
		    g.Bold = False
		    
		    Var riskyMethods() As CodeElement = stats.Value("riskyMethods")
		    Var displayCount As Integer = 0
		    
		    For Each element As CodeElement In riskyMethods
		      If displayCount >= 15 Then  // Limit to 15 for space
		        Exit For element
		      End If
		      
		      For Each pattern As ErrorPattern In element.RiskyPatterns
		        If displayCount >= 15 Then
		          Exit For element
		        End If
		        
		        // Color code by risk level
		        If pattern.RiskLevel = "HIGH" Then
		          g.DrawingColor = Color.RGB(200, 0, 0)
		        Else
		          g.DrawingColor = Color.RGB(255, 140, 0)
		        End If
		        
		        Var riskIcon As String = If(pattern.RiskLevel = "HIGH", "⚠", "⚡")
		        g.DrawText("  " + riskIcon + " " + element.FullPath, margin + 10, yPos)
		        yPos = yPos + lineHeight
		        
		        g.DrawingColor = Color.RGB(100, 100, 100)
		        g.FontSize = 9
		        g.DrawText("    " + pattern.Description, margin + 20, yPos)
		        yPos = yPos + lineHeight + 3
		        
		        g.FontSize = 10
		        displayCount = displayCount + 1
		      Next
		    Next
		    
		    If riskyMethods.Count > 15 Then
		      yPos = yPos + 5
		      g.DrawingColor = Color.RGB(100, 100, 100)
		      g.FontSize = 9
		      Var remaining As Integer = riskyMethods.Count - 15
		      g.DrawText("  ... and " + remaining.ToString + " more methods need error handling", margin + 10, yPos)
		      yPos = yPos + lineHeight
		    End If
		  Else
		    yPos = yPos + 10
		    g.FontSize = 12
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText("✓ All risky operations have error handling!", margin, yPos)
		    yPos = yPos + lineHeight
		  End If
		  
		  yPos = yPos + 20
		  
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
		  g.DrawText("⚠ MOST COMPLEX METHODS (Top 10)", margin, yPos)
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
		    g.DrawText("⚠ Found " + unusedCount + " UNUSED elements:", margin, yPos)
		  Else
		    g.DrawingColor = Color.RGB(0, 150, 0)
		    g.DrawText("✓ No unused elements found", margin, yPos)
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
		  g.DrawText("▼ " + typeName + " (" + count + "):", margin, yPos)
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
