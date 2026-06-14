#tag Class
Protected Class ProjectAnalyzer
	#tag Method, Flags = &h21
		Private Sub AccumulateCodeLine(context As ParsingContext, line As String)
		  //Private Sub AccumulateCodeLine(context As ParsingContext, line As String)
		  // Add line to current method's code if we're in a method
		  
		  If context.InMethodOrFunction Then
		    context.CurrentMethodCode = context.CurrentMethodCode + line + EndOfLine
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeDeepNesting(method As CodeElement, depth As Integer)
		  // Private Sub AnalyzeDeepNesting(method As CodeElement, depth As Integer)
		  Var suggestion As New RefactoringSuggestion(method, "NESTING", "MEDIUM")
		  suggestion.Title = "Deep Nesting (" + depth.ToString + " levels)"
		  suggestion.Description = "Deeply nested code is hard to read and understand."
		  
		  If depth > 5 Then
		    suggestion.Priority = "HIGH"
		  End If
		  
		  suggestion.Suggestions.Add("Use guard clauses with early returns to reduce nesting")
		  suggestion.Suggestions.Add("Extract nested blocks into separate methods")
		  suggestion.Suggestions.Add("Consider inverting conditions to flatten the structure")
		  suggestion.Suggestions.Add("")
		  suggestion.Suggestions.Add("Example refactoring:")
		  suggestion.Suggestions.Add("  Instead of:  If condition1 Then")
		  suggestion.Suggestions.Add("                 If condition2 Then")
		  suggestion.Suggestions.Add("                   // do work")
		  suggestion.Suggestions.Add("  Use:         If Not condition1 Then Return")
		  suggestion.Suggestions.Add("               If Not condition2 Then Return")
		  suggestion.Suggestions.Add("               // do work")
		  
		  method.RefactoringSuggestions.Add(suggestion)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AnalyzeErrorHandling()
		  // Analyze error handling patterns in all methods
		  
		  Logger.Log("=== Starting Error Handling Analysis ===")
		  
		  For Each element As CodeElement In GetMethodElements()  
		    If element.Code.Trim = "" Then Continue
		    
		    // Detect error handling patterns
		    element.HasTryCatch = DetectTryCatchBlocks(element.Code)
		    element.HasDatabaseOperations = DetectDatabaseOperations(element.Code)
		    element.HasFileOperations = DetectFileOperations(element.Code)
		    element.HasNetworkOperations = DetectNetworkOperations(element.Code)
		    element.HasTypeConversions = DetectTypeConversions(element.Code)
		    
		    // Create error patterns for risky operations without error handling
		    If Not element.HasTryCatch Then
		      If element.HasDatabaseOperations Then
		        Var pattern As New ErrorPattern
		        pattern.Element = element
		        pattern.PatternType = "DATABASE"
		        pattern.RiskLevel = "HIGH"
		        pattern.Description = "Database operations without Try/Catch"
		        element.RiskyPatterns.Add(pattern)
		      End If
		      
		      If element.HasFileOperations Then
		        Var pattern As New ErrorPattern
		        pattern.Element = element
		        pattern.PatternType = "FILE_IO"
		        pattern.RiskLevel = "HIGH"
		        pattern.Description = "File operations without Try/Catch"
		        element.RiskyPatterns.Add(pattern)
		      End If
		      
		      If element.HasNetworkOperations Then
		        Var pattern As New ErrorPattern
		        pattern.Element = element
		        pattern.PatternType = "NETWORK"
		        pattern.RiskLevel = "HIGH"
		        pattern.Description = "Network operations without Try/Catch"
		        element.RiskyPatterns.Add(pattern)
		      End If
		      
		      If element.HasTypeConversions Then
		        Var pattern As New ErrorPattern
		        pattern.Element = element
		        pattern.PatternType = "TYPE_CONVERSION"
		        pattern.RiskLevel = "MEDIUM"
		        pattern.Description = "Type conversions without error handling"
		        element.RiskyPatterns.Add(pattern)
		      End If
		    End If
		  Next
		  
		  Logger.Log("=== Error Handling Analysis Complete ===")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeFileForRelationships(content As String)
		  // Track context and find method calls
		  
		  Var lines() As String = content.Split(EndOfLine)
		  Var currentModule As String = ""
		  Var currentClass As String = ""
		  Var currentMethodFullPath As String = ""
		  Var inMethod As Boolean = False
		  
		  For Each line As String In lines
		    line = line.Trim
		    
		    If line.Trim = "" Then Continue
		    
		    ProcessLineForRelationships(line, currentModule, currentClass, _
		    currentMethodFullPath, inMethod)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeHighComplexity(method As CodeElement)
		  // Private Sub AnalyzeHighComplexity(method As CodeElement)
		  Var suggestion As New RefactoringSuggestion(method, "COMPLEXITY", "HIGH")
		  suggestion.Title = "High Cyclomatic Complexity (" + method.CyclomaticComplexity.ToString + ")"
		  suggestion.Description = "This method has high complexity, making it difficult to understand and test."
		  
		  // Count conditional statements
		  Var ifCount As Integer = CountOccurrencesInString(method.Code.Uppercase, " IF ")
		  Var caseCount As Integer = CountOccurrencesInString(method.Code.Uppercase, " CASE ")
		  
		  If ifCount > 5 Then
		    suggestion.Suggestions.Add("Extract " + ifCount.ToString + " conditional checks into separate validation methods")
		    suggestion.Suggestions.Add("Use guard clauses (early returns) to reduce nesting")
		  End If
		  
		  If caseCount > 0 Then
		    suggestion.Suggestions.Add("Consider using a Dictionary or Strategy pattern instead of Select/Case")
		  End If
		  
		  If method.CyclomaticComplexity > 15 Then
		    suggestion.Suggestions.Add("URGENT: Split this method into smaller, focused methods")
		    suggestion.Suggestions.Add("Target complexity: < 10 per method")
		  Else
		    suggestion.Suggestions.Add("Refactor to reduce complexity below 10")
		  End If
		  
		  method.RefactoringSuggestions.Add(suggestion)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeLongMethod(method As CodeElement)
		  // Private Sub AnalyzeLongMethod(method As CodeElement)
		  Var suggestion As New RefactoringSuggestion(method, "LENGTH", "MEDIUM")
		  suggestion.Title = "Long Method (" + method.LinesOfCode.ToString + " lines)"
		  suggestion.Description = "This method is too long, making it hard to understand its purpose."
		  
		  If method.LinesOfCode > 100 Then
		    suggestion.Priority = "HIGH"
		    suggestion.Suggestions.Add("URGENT: This method is extremely long - split into multiple methods")
		  End If
		  
		  suggestion.Suggestions.Add("Identify distinct logical sections and extract them into separate methods")
		  suggestion.Suggestions.Add("Target: Keep methods under 30 lines for better readability")
		  suggestion.Suggestions.Add("Look for repeated code blocks that can be extracted")
		  
		  // Check for comments that indicate sections
		  If method.Code.Contains("// ") Or method.Code.Contains("' ") Then
		    suggestion.Suggestions.Add("Your comments suggest natural break points - each commented section could be a method")
		  End If
		  
		  method.RefactoringSuggestions.Add(suggestion)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeManyParameters(method As CodeElement)
		  // Private Sub AnalyzeManyParameters(method As CodeElement)
		  Var suggestion As New RefactoringSuggestion(method, "PARAMETERS", "MEDIUM")
		  suggestion.Title = "Too Many Parameters (" + method.ParameterCount.ToString + ")"
		  suggestion.Description = "Methods with many parameters are hard to call and maintain."
		  
		  // Count ByRef parameters
		  Var byRefCount As Integer = CountOccurrencesInString(method.Parameters.Uppercase, "BYREF")
		  
		  If byRefCount > 3 Then
		    suggestion.Priority = "HIGH"
		    suggestion.Suggestions.Add("CRITICAL: " + byRefCount.ToString + " ByRef parameters detected - consider returning a result object instead")
		    suggestion.Suggestions.Add("Create a Result class to return multiple values")
		  End If
		  
		  suggestion.Suggestions.Add("Create a parameter object to group related parameters")
		  suggestion.Suggestions.Add("Example: Instead of DrawNode(x, y, width, height, color, label)")
		  suggestion.Suggestions.Add("         Use: DrawNode(nodeConfig As NodeConfiguration)")
		  
		  If method.ParameterCount > 7 Then
		    suggestion.Suggestions.Add("Consider using a Builder pattern for complex object construction")
		  End If
		  
		  method.RefactoringSuggestions.Add(suggestion)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AnalyzeMissingErrorHandling(method As CodeElement)
		  // Private Sub AnalyzeMissingErrorHandling(method As CodeElement)
		  
		  Var suggestion As New RefactoringSuggestion(method, "ERROR_HANDLING", "HIGH")
		  suggestion.Title = "Missing Error Handling"
		  suggestion.Description = "This method performs risky operations without proper error handling."
		  
		  // Build list of risky operations
		  Var riskyOps() As String
		  For Each pattern As ErrorPattern In method.RiskyPatterns
		    riskyOps.Add(pattern.PatternType + " (" + pattern.RiskLevel + ")")
		  Next
		  
		  // Manually join the array
		  Var joinedOps As String = ""
		  For i As Integer = 0 To riskyOps.LastIndex
		    If i > 0 Then joinedOps = joinedOps + ", "
		    joinedOps = joinedOps + riskyOps(i)
		  Next
		  suggestion.Suggestions.Add("Add Try/Catch block to handle: " + joinedOps)
		  
		  
		  // Provide specific examples based on risk type
		  For Each pattern As ErrorPattern In method.RiskyPatterns
		    Select Case pattern.PatternType
		    Case "DATABASE"
		      suggestion.Suggestions.Add("Catch DatabaseException for SQL errors")
		    Case "FILE_IO"
		      suggestion.Suggestions.Add("Catch IOException for file access errors")
		    Case "NETWORK"
		      suggestion.Suggestions.Add("Catch SocketException or IOException for network errors")
		    Case "TYPE_CONVERSION"
		      suggestion.Suggestions.Add("Validate data before conversion or use Try/Catch")
		    End Select
		  Next
		  
		  // Show example code structure
		  suggestion.Suggestions.Add("")
		  suggestion.Suggestions.Add("Example structure:")
		  suggestion.Suggestions.Add("  Try")
		  suggestion.Suggestions.Add("    // Your risky operation here")
		  suggestion.Suggestions.Add("  Catch e As IOException")
		  suggestion.Suggestions.Add("    // Handle the error appropriately")
		  suggestion.Suggestions.Add("  End Try")
		  
		  method.RefactoringSuggestions.Add(suggestion)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AnalyzeRefactoringOpportunities()
		  // Public Sub AnalyzeRefactoringOpportunities()
		  Logger.Log("=== Starting Refactoring Analysis ===")
		  
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    // Skip if no code
		    If method.Code.Trim = "" Then Continue
		    
		    // Calculate LOC if not already done
		    If method.LinesOfCode = 0 Then
		      call method.CalculateLinesOfCode()
		    End If
		    
		    // Analyze complexity
		    If method.CyclomaticComplexity > 10 Then
		      AnalyzeHighComplexity(method)
		    End If
		    
		    // Analyze length
		    If method.LinesOfCode > 50 Then
		      AnalyzeLongMethod(method)
		    End If
		    
		    // Analyze parameters
		    If method.ParameterCount > 5 Then
		      AnalyzeManyParameters(method)
		    End If
		    
		    // Analyze error handling
		    If method.RiskyPatterns.Count > 0 Then
		      AnalyzeMissingErrorHandling(method)
		    End If
		    
		    // Analyze nesting depth
		    Var nestingDepth As Integer = CalculateNestingDepth(method.Code)
		    If nestingDepth > 3 Then
		      AnalyzeDeepNesting(method, nestingDepth)
		    End If
		  Next
		  
		  Logger.Log("=== Refactoring Analysis Complete ===")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub BuildDependencyData(ByRef nodes() As MethodNode, ByRef deps() As Dependency)
		  // Public Sub BuildDependencyData(ByRef nodes() As MethodNode, ByRef deps() As Dependency)
		  // Build nodes from methods
		  For Each element As CodeElement In mElements
		    If element.ElementType = "METHOD" Then
		      Var node As New MethodNode
		      node.Name = element.Name
		      node.FullPath = element.FullPath
		      node.Complexity = element.CyclomaticComplexity
		      node.LOC = element.LinesOfCode
		      node.CalledBy = 0
		      
		      // Color by complexity
		      If node.Complexity <= 5 Then
		        node.NodeColor = Color.RGB(46, 204, 113) // Green
		      ElseIf node.Complexity <= 10 Then
		        node.NodeColor = Color.RGB(241, 196, 15) // Yellow
		      Else
		        node.NodeColor = Color.RGB(231, 76, 60) // Red
		      End If
		      
		      nodes.Add(node)
		    End If
		  Next
		  
		  // Build dependencies from code analysis
		  // Look through each element's code for method calls
		  For Each element As CodeElement In mElements
		    If element.ElementType = "METHOD" And element.Code.Trim <> "" Then
		      // Scan the code for calls to other methods
		      For Each targetElement As CodeElement In mElements
		        If targetElement.ElementType = "METHOD" And targetElement <> element Then
		          // Check if this method is called in the code
		          If element.Code.IndexOf(targetElement.Name) > -1 Then
		            Var dep As New Dependency
		            dep.FromPath = element.FullPath
		            dep.ToPath = targetElement.FullPath
		            deps.Add(dep)
		          End If
		        End If
		      Next
		    End If
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildFullPath(moduleName As String, className As String, elementName As String) As String
		  // Private Function BuildFullPath(moduleName As String, className As String, elementName As String) As String
		  If moduleName <> "" And className <> "" Then
		    Return moduleName + "." + className + "." + elementName
		  ElseIf moduleName <> "" Then
		    Return moduleName + "." + elementName
		  ElseIf className <> "" Then
		    Return className + "." + elementName
		  Else
		    Return elementName
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildInteractiveHTML(nodesJSON As String, edgesJSON As String) As String
		  // Private Function BuildInteractiveHTML(nodesJSON As String, edgesJSON As String) As String
		  Var html As String = "<!DOCTYPE html>" + EndOfLine
		  html = html + "<html>" + EndOfLine
		  html = html + "<head>" + EndOfLine
		  html = html + "  <meta charset=""utf-8"">" + EndOfLine
		  html = html + "  <title>Interactive Dependency Graph</title>" + EndOfLine
		  html = html + "  <script type=""text/javascript"" src=""https://unpkg.com/vis-network@9.1.6/standalone/umd/vis-network.min.js""></script>" + EndOfLine
		  
		  html = html + "  <style>" + EndOfLine
		  html = html + "    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }" + EndOfLine
		  html = html + "    h1 { margin-top: 0; }" + EndOfLine
		  html = html + "    #controls { background: white; padding: 15px; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); position: sticky; top: 0; z-index: 100; }" + EndOfLine
		  html = html + "    #mynetwork { width: 100%; height: 90vh; border: 1px solid #ddd; background: white; border-radius: 8px; min-height: 800px; }" + EndOfLine
		  html = html + "    .control-group { display: inline-block; margin-right: 20px; margin-bottom: 10px; }" + EndOfLine
		  html = html + "    label { font-weight: bold; margin-right: 8px; }" + EndOfLine
		  html = html + "    input, select, button { padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; }" + EndOfLine
		  html = html + "    button { background: #3498db; color: white; cursor: pointer; border: none; }" + EndOfLine
		  html = html + "    button:hover { background: #2980b9; }" + EndOfLine
		  html = html + "    .legend { margin-top: 10px; }" + EndOfLine
		  html = html + "    .legend-item { display: inline-block; margin-right: 20px; }" + EndOfLine
		  html = html + "    .legend-color { display: inline-block; width: 20px; height: 20px; border-radius: 50%; margin-right: 5px; vertical-align: middle; }" + EndOfLine
		  html = html + "    #info { background: #ecf0f1; padding: 10px; margin-top: 10px; border-radius: 4px; display: none; }" + EndOfLine
		  html = html + "    #error { background: #e74c3c; color: white; padding: 10px; margin-bottom: 10px; border-radius: 4px; display: none; }" + EndOfLine
		  html = html + "  </style>" + EndOfLine
		  html = html + "</head>" + EndOfLine
		  html = html + "<body>" + EndOfLine
		  html = html + "  <h1>Interactive Method Dependency Graph</h1>" + EndOfLine
		  html = html + "  <div id=""error""></div>" + EndOfLine
		  html = html + "  <div id=""controls"">" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <label>Search:</label>" + EndOfLine
		  html = html + "      <input type=""text"" id=""searchBox"" placeholder=""Method name..."" />" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <label>Complexity:</label>" + EndOfLine
		  html = html + "      <select id=""complexityFilter"">" + EndOfLine
		  html = html + "        <option value=""all"" selected>All</option>" + EndOfLine
		  html = html + "        <option value=""low"">Low (1-5)</option>" + EndOfLine
		  html = html + "        <option value=""medium"">Medium (6-10)</option>" + EndOfLine
		  html = html + "        <option value=""high"">High (11+)</option>" + EndOfLine
		  html = html + "      </select>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <button onclick=""resetView()"">Reset View</button>" + EndOfLine
		  html = html + "      <button onclick=""fitNetwork()"">Fit All</button>" + EndOfLine
		  html = html + "      <button onclick=""zoomIn()"">Zoom In</button>" + EndOfLine
		  html = html + "      <button onclick=""zoomOut()"">Zoom Out</button>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""legend"">" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#2ecc71""></span>Low Complexity (1-5)</span>" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#f1c40f""></span>Medium Complexity (6-10)</span>" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#e74c3c""></span>High Complexity (11+)</span>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div id=""info""></div>" + EndOfLine
		  html = html + "  </div>" + EndOfLine
		  html = html + "  <div id=""mynetwork""></div>" + EndOfLine
		  
		  html = html + "  <script type=""text/javascript"">" + EndOfLine
		  html = html + "    window.onerror = function(msg, url, line) {" + EndOfLine
		  html = html + "      document.getElementById('error').innerHTML = 'Error: ' + msg + ' (line ' + line + ')';" + EndOfLine
		  html = html + "      document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "      return false;" + EndOfLine
		  html = html + "    };" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    if (typeof vis === 'undefined') {" + EndOfLine
		  html = html + "      document.getElementById('error').innerHTML = 'Failed to load vis.js library. Check internet connection.';" + EndOfLine
		  html = html + "      document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "    } else {" + EndOfLine
		  html = html + "      initializeGraph();" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    var allNodes, allEdges, network;" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function initializeGraph() {" + EndOfLine
		  html = html + "      try {" + EndOfLine
		  html = html + "        allNodes = " + nodesJSON + ";" + EndOfLine
		  html = html + "        allEdges = " + edgesJSON + ";" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        console.Logger.Log('Loaded ' + allNodes.length + ' nodes and ' + allEdges.length + ' edges');" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        var nodes = new vis.DataSet(allNodes);" + EndOfLine
		  html = html + "        var edges = new vis.DataSet(allEdges);" + EndOfLine
		  html = html + "        var container = document.getElementById('mynetwork');" + EndOfLine
		  html = html + "        var data = { nodes: nodes, edges: edges };" + EndOfLine
		  html = html + "        var options = {" + EndOfLine
		  html = html + "          physics: {" + EndOfLine
		  html = html + "            enabled: true," + EndOfLine
		  html = html + "            barnesHut: {" + EndOfLine
		  html = html + "              gravitationalConstant: -5000," + EndOfLine
		  html = html + "              centralGravity: 0.05," + EndOfLine
		  html = html + "              springLength: 250," + EndOfLine
		  html = html + "              springConstant: 0.02," + EndOfLine
		  html = html + "              damping: 0.95," + EndOfLine
		  html = html + "              avoidOverlap: 1" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            stabilization: { iterations: 300 }" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          layout: { randomSeed: 42, improvedLayout: true }," + EndOfLine
		  html = html + "          interaction: {" + EndOfLine
		  html = html + "            hover: true," + EndOfLine
		  html = html + "            tooltipDelay: 100," + EndOfLine
		  html = html + "            zoomView: true," + EndOfLine
		  html = html + "            dragView: true," + EndOfLine
		  html = html + "            navigationButtons: true," + EndOfLine
		  html = html + "            keyboard: true" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          nodes: {" + EndOfLine
		  html = html + "            shape: 'box'," + EndOfLine
		  html = html + "            borderWidth: 2," + EndOfLine
		  html = html + "            borderWidthSelected: 3," + EndOfLine
		  html = html + "            margin: 10" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          edges: {" + EndOfLine
		  html = html + "            width: 2," + EndOfLine
		  html = html + "            arrows: { to: { enabled: true, scaleFactor: 0.5 } }," + EndOfLine
		  html = html + "            smooth: { type: 'cubicBezier', forceDirection: 'horizontal' }" + EndOfLine
		  html = html + "          }" + EndOfLine
		  html = html + "        };" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network = new vis.Network(container, data, options);" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network.once('stabilizationIterationsDone', function() {" + EndOfLine
		  html = html + "          network.setOptions({ physics: false });" + EndOfLine
		  html = html + "          setTimeout(function() {" + EndOfLine
		  html = html + "            network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "          }, 100);" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network.on('click', function(params) {" + EndOfLine
		  html = html + "          if (params.nodes.length > 0) {" + EndOfLine
		  html = html + "            highlightConnections(params.nodes[0]);" + EndOfLine
		  html = html + "            showNodeInfo(params.nodes[0]);" + EndOfLine
		  html = html + "          } else {" + EndOfLine
		  html = html + "            resetHighlight();" + EndOfLine
		  html = html + "            document.getElementById('info').style.display = 'none';" + EndOfLine
		  html = html + "          }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        setupControls();" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "      } catch(e) {" + EndOfLine
		  html = html + "        document.getElementById('error').innerHTML = 'Error initializing: ' + e.message;" + EndOfLine
		  html = html + "        document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "        console.error(e);" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function highlightConnections(nodeId) {" + EndOfLine
		  html = html + "      var connectedNodes = network.getConnectedNodes(nodeId);" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      allNodes.forEach(function(node) {" + EndOfLine
		  html = html + "        if (node.id === nodeId) {" + EndOfLine
		  html = html + "          // CLICKED NODE - bright blue with thick border" + EndOfLine
		  html = html + "          network.body.data.nodes.update({" + EndOfLine
		  html = html + "            id: node.id," + EndOfLine
		  html = html + "            opacity: 1," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            borderWidth: 5," + EndOfLine
		  html = html + "            color: {" + EndOfLine
		  html = html + "              background: '#3498db'," + EndOfLine
		  html = html + "              border: '#000000'," + EndOfLine
		  html = html + "              highlight: { background: '#3498db', border: '#000000' }" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            font: { color: 'white', size: 14, bold: true }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else if (connectedNodes.indexOf(node.id) !== -1) {" + EndOfLine
		  html = html + "          // CONNECTED NODES - keep original color but with bright border" + EndOfLine
		  html = html + "          var originalNode = allNodes.find(n => n.id === node.id);" + EndOfLine
		  html = html + "          var originalColor = originalNode.color.background || originalNode.color;" + EndOfLine
		  html = html + "          network.body.data.nodes.update({" + EndOfLine
		  html = html + "            id: node.id," + EndOfLine
		  html = html + "            opacity: 1," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            borderWidth: 4," + EndOfLine
		  html = html + "            color: {" + EndOfLine
		  html = html + "              background: originalColor," + EndOfLine
		  html = html + "              border: '#FFD700'," + EndOfLine
		  html = html + "              highlight: { background: originalColor, border: '#FFD700' }" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            font: { color: 'white', size: 12, bold: true }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // HIDE unconnected nodes" + EndOfLine
		  html = html + "          network.body.data.nodes.update({id: node.id, hidden: true});" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      // Update edges - make connected edges bright and thick" + EndOfLine
		  html = html + "      allEdges.forEach(function(edge) {" + EndOfLine
		  html = html + "        if (edge.from === nodeId || edge.to === nodeId) {" + EndOfLine
		  html = html + "          // Edges directly from/to clicked node - BRIGHT BLUE & THICK" + EndOfLine
		  html = html + "          network.body.data.edges.update({" + EndOfLine
		  html = html + "            from: edge.from," + EndOfLine
		  html = html + "            to: edge.to," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            width: 4," + EndOfLine
		  html = html + "            color: { color: '#3498db', highlight: '#3498db' }," + EndOfLine
		  html = html + "            arrows: { to: { enabled: true, scaleFactor: 1 } }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // Hide other edges" + EndOfLine
		  html = html + "          network.body.data.edges.update({from: edge.from, to: edge.to, hidden: true});" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      // Auto-fit to show the subgraph" + EndOfLine
		  html = html + "      setTimeout(function() {" + EndOfLine
		  html = html + "        network.fit({" + EndOfLine
		  html = html + "          nodes: [nodeId].concat(connectedNodes)," + EndOfLine
		  html = html + "          padding: 100," + EndOfLine
		  html = html + "          animation: { duration: 400 }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      }, 100);" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function resetHighlight() {" + EndOfLine
		  html = html + "      allNodes.forEach(function(node) {" + EndOfLine
		  html = html + "        var originalColor = node.color.background || node.color;" + EndOfLine
		  html = html + "        network.body.data.nodes.update({" + EndOfLine
		  html = html + "          id: node.id," + EndOfLine
		  html = html + "          opacity: 1," + EndOfLine
		  html = html + "          hidden: false," + EndOfLine
		  html = html + "          borderWidth: 2," + EndOfLine
		  html = html + "          color: {" + EndOfLine
		  html = html + "            background: originalColor," + EndOfLine
		  html = html + "            border: '#2c3e50'," + EndOfLine
		  html = html + "            highlight: { background: originalColor, border: '#000000' }" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          font: { color: 'white', size: 10 }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      allEdges.forEach(function(edge) {" + EndOfLine
		  html = html + "        network.body.data.edges.update({" + EndOfLine
		  html = html + "          from: edge.from," + EndOfLine
		  html = html + "          to: edge.to," + EndOfLine
		  html = html + "          hidden: false," + EndOfLine
		  html = html + "          width: 2," + EndOfLine
		  html = html + "          color: { color: '#95a5a6', highlight: '#3498db' }," + EndOfLine
		  html = html + "          arrows: { to: { enabled: true, scaleFactor: 0.5 } }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function showNodeInfo(nodeId) {" + EndOfLine
		  html = html + "      var node = allNodes.find(n => n.id === nodeId);" + EndOfLine
		  html = html + "      if (node) {" + EndOfLine
		  html = html + "        var info = document.getElementById('info');" + EndOfLine
		  html = html + "        info.innerHTML = '<strong>' + node.label + '</strong><br>Complexity: ' + node.complexity + ' | LOC: ' + node.loc;" + EndOfLine
		  html = html + "        info.style.display = 'block';" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function setupControls() {" + EndOfLine
		  html = html + "      document.getElementById('searchBox').addEventListener('input', function(e) {" + EndOfLine
		  html = html + "        var term = e.target.value.toLowerCase();" + EndOfLine
		  html = html + "        if (!term) {" + EndOfLine
		  html = html + "          resetHighlight();" + EndOfLine
		  html = html + "          return;" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "        var matches = allNodes.filter(n => n.label.toLowerCase().includes(term));" + EndOfLine
		  html = html + "        if (matches.length > 0) {" + EndOfLine
		  html = html + "          network.selectNodes([matches[0].id]);" + EndOfLine
		  html = html + "          network.focus(matches[0].id, { scale: 1.2, animation: { duration: 500 } });" + EndOfLine
		  html = html + "          highlightConnections(matches[0].id);" + EndOfLine
		  html = html + "          showNodeInfo(matches[0].id);" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "      document.getElementById('complexityFilter').addEventListener('change', function(e) {" + EndOfLine
		  html = html + "        var filter = e.target.value;" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        if (filter === 'all') {" + EndOfLine
		  html = html + "          // Restore everything" + EndOfLine
		  html = html + "          network.body.data.nodes.clear();" + EndOfLine
		  html = html + "          network.body.data.edges.clear();" + EndOfLine
		  html = html + "          network.body.data.nodes.add(allNodes);" + EndOfLine
		  html = html + "          network.body.data.edges.add(allEdges);" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // Filter nodes" + EndOfLine
		  html = html + "          var filtered = allNodes.filter(function(node) {" + EndOfLine
		  html = html + "            if (filter === 'low') return node.complexity <= 5;" + EndOfLine
		  html = html + "            if (filter === 'medium') return node.complexity > 5 && node.complexity <= 10;" + EndOfLine
		  html = html + "            if (filter === 'high') return node.complexity > 10;" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          // Get IDs of filtered nodes" + EndOfLine
		  html = html + "          var filteredIds = filtered.map(n => n.id);" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          // Filter edges to only show connections between filtered nodes" + EndOfLine
		  html = html + "          var filteredEdges = allEdges.filter(function(edge) {" + EndOfLine
		  html = html + "            return filteredIds.indexOf(edge.from) !== -1 && filteredIds.indexOf(edge.to) !== -1;" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          network.body.data.nodes.clear();" + EndOfLine
		  html = html + "          network.body.data.edges.clear();" + EndOfLine
		  html = html + "          network.body.data.nodes.add(filtered);" + EndOfLine
		  html = html + "          network.body.data.edges.add(filteredEdges);" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        setTimeout(function() {" + EndOfLine
		  html = html + "          network.fit({ padding: 100, animation: { duration: 500 } });" + EndOfLine
		  html = html + "        }, 50);" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function resetView() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        // Clear and restore both nodes and edges" + EndOfLine
		  html = html + "        network.body.data.nodes.clear();" + EndOfLine
		  html = html + "        network.body.data.edges.clear();" + EndOfLine
		  html = html + "        network.body.data.nodes.add(allNodes);" + EndOfLine
		  html = html + "        network.body.data.edges.add(allEdges);" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Reset all styling" + EndOfLine
		  html = html + "        resetHighlight();" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Clear search and filter" + EndOfLine
		  html = html + "        document.getElementById('searchBox').value = '';" + EndOfLine
		  html = html + "        document.getElementById('complexityFilter').value = 'all';" + EndOfLine
		  html = html + "        document.getElementById('info').style.display = 'none';" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Re-enable physics briefly to reorganize" + EndOfLine
		  html = html + "        network.setOptions({ physics: { enabled: true } });" + EndOfLine
		  html = html + "        setTimeout(function() {" + EndOfLine
		  html = html + "          network.setOptions({ physics: { enabled: false } });" + EndOfLine
		  html = html + "          network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "        }, 500);" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function fitNetwork() {" + EndOfLine
		  html = html + "      if (network) network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function zoomIn() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        var scale = network.getScale();" + EndOfLine
		  html = html + "        network.moveTo({ scale: scale * 1.2, animation: { duration: 300 } });" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function zoomOut() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        var scale = network.getScale();" + EndOfLine
		  html = html + "        network.moveTo({ scale: scale * 0.8, animation: { duration: 300 } });" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "  </script>" + EndOfLine
		  html = html + "</body>" + EndOfLine
		  html = html + "</html>" + EndOfLine
		  
		  Return html
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function BuildMethodFullPath(line As String, currentModule As String, currentClass As String) As String
		  // Build the full path for a method based on context
		  
		  Var methodName As String = ExtractMethodName(line)
		  If methodName.Trim = "" Then Return ""
		  
		  If currentModule.Trim <> "" And currentClass.Trim <> "" Then
		    Return currentModule + "." + currentClass + "." + methodName
		  ElseIf currentModule.Trim <> "" Then
		    Return currentModule + "." + methodName
		  ElseIf currentClass.Trim <> "" Then
		    Return currentClass + "." + methodName
		  Else
		    Return methodName
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub BuildRelationships(projectFolder As FolderItem)
		  // Scan all files again to find relationships between elements
		  ScanForRelationships()
		  
		  Logger.Log("=== BuildRelationships STARTED ===")
		  // Build parent-child relationships
		  For Each element As CodeElement In mElements
		    If element.ElementType = "METHOD" Or element.ElementType = "PROPERTY" Or element.ElementType = "VARIABLE" Then
		      // Find parent class or module
		      Var parentPath As String = ""
		      If element.FullPath.Contains(".") Then
		        Var parts() As String = element.FullPath.Split(".")
		        If parts.Count >= 2 Then
		          // Get everything except the last part
		          For i As Integer = 0 To parts.Count - 2
		            If parentPath.Trim <> "" Then
		              parentPath = parentPath + "."
		            End If
		            parentPath = parentPath + parts(i)
		          Next
		          
		          // Find the parent element
		          If ElementLookup.HasKey(parentPath) Then
		            Var parentElement As CodeElement = ElementLookup.Value(parentPath)
		            element.Parent = parentElement
		            parentElement.Contains.Add(element)
		          End If
		        End If
		      End If
		    End If
		  Next
		  
		  // Scan all files for relationships
		  ScanForRelationships()
		  
		  // Check what we found
		  Var allElements() As CodeElement = GetAllElements()
		  Var methodsWithCalls As Integer = 0
		  Var totalCalls As Integer = 0
		  
		  For Each element As CodeElement In allElements
		    If element.ElementType = "METHOD" Then
		      If element.CallsTo.Count > 0 Then
		        methodsWithCalls = methodsWithCalls + 1
		        totalCalls = totalCalls + element.CallsTo.Count
		        Logger.Log("Method " + element.Name + " calls " + element.CallsTo.Count.ToString + " other methods")
		      End If
		    End If
		  Next
		  
		  Var methodsStr As String = methodsWithCalls.ToString
		  Var callsStr As String = totalCalls.ToString
		  Logger.Log("Methods with calls: " + methodsStr)
		  Logger.Log("Total calls: " + callsStr)
		  Logger.Log("=== BuildRelationships COMPLETED ===")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CalculateGraphLayout(nodes() As MethodNode, deps() As Dependency, width As Integer, height As Integer)
		  ///Public Sub CalculateGraphLayout(nodes() As MethodNode, deps() As Dependency, width As Integer, height As Integer)
		  If nodes.Count = 0 Then Return
		  
		  // Group by class/module
		  Var groups As New Dictionary
		  
		  For Each node As MethodNode In nodes
		    Var parts() As String = node.FullPath.Split(".")
		    Var className As String
		    
		    If parts.Count >= 2 Then
		      className = parts(0)
		    Else
		      className = "Global"
		    End If
		    
		    If Not groups.HasKey(className) Then
		      Var groupArray() As MethodNode
		      groups.Value(className) = groupArray
		    End If
		    
		    Var groupArray() As MethodNode = groups.Value(className)
		    groupArray.Add(node)
		    groups.Value(className) = groupArray
		  Next
		  
		  // Layout parameters - MUCH MORE HORIZONTAL SPACE
		  Var startY As Integer = 120
		  Var rowHeight As Integer = 120  // More vertical space between rows
		  Var currentY As Integer = startY
		  Var nodesPerRow As Integer = 8  // More nodes per row (was 6)
		  Var nodeSpacing As Integer = 180  // More horizontal space (was 130)
		  
		  For Each className As Variant In groups.Keys
		    Var groupNodes() As MethodNode = groups.Value(className)
		    
		    // Calculate how many rows needed for this group
		    Var rows As Integer = Ceil(groupNodes.Count / nodesPerRow)
		    
		    For rowIndex As Integer = 0 To rows - 1
		      Var startIdx As Integer = rowIndex * nodesPerRow
		      Var endIdx As Integer = Min(startIdx + nodesPerRow - 1, groupNodes.Count - 1)
		      Var nodesInRow As Integer = endIdx - startIdx + 1
		      
		      // Center the row
		      Var totalRowWidth As Integer = nodesInRow * nodeSpacing
		      Var startX As Integer = (width - totalRowWidth) \ 2
		      
		      // Position each node in the row
		      For i As Integer = startIdx To endIdx
		        Var posInRow As Integer = i - startIdx
		        groupNodes(i).X = startX + (posInRow * nodeSpacing) + nodeSpacing \ 2
		        groupNodes(i).Y = currentY
		      Next
		      
		      currentY = currentY + rowHeight
		    Next
		    
		    // Extra space between groups
		    currentY = currentY + 30
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateMethodComplexity(method As CodeElement) As Integer
		  // Private Function CalculateMethodComplexity(method As CodeElement) As Integer
		  ' Calculate cyclomatic complexity for a method
		  
		  Var complexity As Integer = 1
		  Var upperCode As String = method.Code.Uppercase
		  
		  ' Decision points
		  complexity = complexity + CountOccurrencesInString(upperCode, " IF ")
		  complexity = complexity + CountOccurrencesInString(upperCode, EndOfLine + "IF ")
		  complexity = complexity + CountOccurrencesInString(upperCode, "ELSEIF ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " FOR ")
		  complexity = complexity + CountOccurrencesInString(upperCode, EndOfLine + "FOR ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " WHILE ")
		  complexity = complexity + CountOccurrencesInString(upperCode, EndOfLine + "WHILE ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " DO ")
		  complexity = complexity + CountOccurrencesInString(upperCode, EndOfLine + "DO ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " CASE ")
		  complexity = complexity + CountOccurrencesInString(upperCode, EndOfLine + "CASE ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " AND ")
		  complexity = complexity + CountOccurrencesInString(upperCode, " OR ")
		  
		  Return complexity
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CalculateNestingDepth(code As String) As Integer
		  // Private Function CalculateNestingDepth(code As String) As Integer
		  Var lines() As String = code.Split(EndOfLine)
		  Var maxDepth As Integer = 0
		  Var currentDepth As Integer = 0
		  
		  For Each line As String In lines
		    Var trimmed As String = line.Trim.Uppercase
		    
		    // Increase depth for block starts
		    If trimmed.BeginsWith("IF ") Or trimmed.BeginsWith("FOR ") Or _
		      trimmed.BeginsWith("WHILE ") Or trimmed.BeginsWith("SELECT ") Or _
		      trimmed.BeginsWith("TRY") Then
		      currentDepth = currentDepth + 1
		      maxDepth = Max(maxDepth, currentDepth)
		    End If
		    
		    // Decrease depth for block ends
		    If trimmed.BeginsWith("END IF") Or trimmed.BeginsWith("NEXT") Or _
		      trimmed.BeginsWith("WEND") Or trimmed.BeginsWith("END SELECT") Or _
		      trimmed.BeginsWith("END TRY") Then
		      currentDepth = currentDepth - 1
		    End If
		  Next
		  
		  Return maxDepth
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CalculateQualityScore() As QualityScore
		  //Function CalculateQualityScore() As QualityScore
		  ' Calculate comprehensive quality score for the project
		  
		  Var score As New QualityScore
		  
		  Var allMethods() As CodeElement = GetMethodElements()
		  
		  If allMethods.Count = 0 Then
		    ' No methods to analyze
		    score.OverallScore = 0
		    score.Grade = "N/A"
		    Return score
		  End If
		  
		  ' ==============================================
		  ' 1. ERROR HANDLING COVERAGE (30%)
		  ' ==============================================
		  Var methodsWithRiskyOps As Integer = 0
		  Var methodsWithErrorHandling As Integer = 0
		  
		  For Each method As CodeElement In allMethods
		    ' Check if method has risky operations
		    Var hasRiskyOps As Boolean = False
		    Var code As String = method.Code.Uppercase
		    
		    ' Database operations
		    If code.IndexOf("DATABASE") >= 0 Or code.IndexOf("RECORDSET") >= 0 Or _
		      code.IndexOf("SQL") >= 0 Or code.IndexOf("QUERY") >= 0 Then
		      hasRiskyOps = True
		    End If
		    
		    ' File operations
		    If code.IndexOf("FOLDERITEM") >= 0 Or code.IndexOf("TEXTINPUTSTREAM") >= 0 Or _
		      code.IndexOf("TEXTOUTPUTSTREAM") >= 0 Or code.IndexOf("BINARYSTREAM") >= 0 Then
		      hasRiskyOps = True
		    End If
		    
		    ' Network operations
		    If code.IndexOf("URLCONNECTION") >= 0 Or code.IndexOf("SOCKET") >= 0 Or _
		      code.IndexOf("HTTP") >= 0 Then
		      hasRiskyOps = True
		    End If
		    
		    ' Type conversions
		    If code.IndexOf(".TOINTEGER") >= 0 Or code.IndexOf(".TODOUBLE") >= 0 Or _
		      code.IndexOf("VAL(") >= 0 Or code.IndexOf("CTYPE(") >= 0 Then
		      hasRiskyOps = True
		    End If
		    
		    If hasRiskyOps Then
		      methodsWithRiskyOps = methodsWithRiskyOps + 1
		      
		      ' Check if it has error handling
		      If code.IndexOf("TRY") >= 0 And code.IndexOf("CATCH") >= 0 Then
		        methodsWithErrorHandling = methodsWithErrorHandling + 1
		      End If
		    End If
		  Next
		  
		  If methodsWithRiskyOps > 0 Then
		    score.ErrorHandlingCoverage = (methodsWithErrorHandling / methodsWithRiskyOps) * 100
		    score.ErrorHandlingScore = score.ErrorHandlingCoverage
		  Else
		    ' No risky operations found - perfect score
		    score.ErrorHandlingCoverage = 100
		    score.ErrorHandlingScore = 100
		  End If
		  
		  ' ==============================================
		  ' 2. AVERAGE COMPLEXITY (25%)
		  ' ==============================================
		  Var totalComplexity As Integer = 0
		  Var methodsWithCode As Integer = 0
		  
		  For Each method As CodeElement In allMethods
		    If method.Code.Trim <> "" Then
		      Var complexity As Integer = CalculateMethodComplexity(method)
		      totalComplexity = totalComplexity + complexity
		      methodsWithCode = methodsWithCode + 1
		    End If
		  Next
		  
		  If methodsWithCode > 0 Then
		    score.AverageComplexity = totalComplexity / methodsWithCode
		  Else
		    score.AverageComplexity = 0
		  End If
		  
		  ' Score complexity (lower is better)
		  ' Excellent: 1-10 = 100
		  ' Good: 11-15 = 80
		  ' Fair: 16-20 = 60
		  ' Poor: 21-30 = 40
		  ' Critical: 31+ = 20
		  If score.AverageComplexity <= 10 Then
		    score.ComplexityScore = 100
		  ElseIf score.AverageComplexity <= 15 Then
		    score.ComplexityScore = 100 - ((score.AverageComplexity - 10) * 4)
		  ElseIf score.AverageComplexity <= 20 Then
		    score.ComplexityScore = 80 - ((score.AverageComplexity - 15) * 4)
		  ElseIf score.AverageComplexity <= 30 Then
		    score.ComplexityScore = 60 - ((score.AverageComplexity - 20) * 2)
		  Else
		    score.ComplexityScore = 20
		  End If
		  
		  ' ==============================================
		  ' 3. CODE REUSE (LOW UNUSED CODE) (20%)
		  ' ==============================================
		  Var allElements() As CodeElement = GetAllElements()
		  Var unusedElements() As CodeElement = GetUnusedElements()
		  
		  If allElements.Count > 0 Then
		    score.UnusedPercentage = (unusedElements.Count / allElements.Count) * 100
		    score.CodeReuseScore = 100 - score.UnusedPercentage
		    
		    ' Ensure score doesn't go negative
		    If score.CodeReuseScore < 0 Then
		      score.CodeReuseScore = 0
		    End If
		  Else
		    score.UnusedPercentage = 0
		    score.CodeReuseScore = 100
		  End If
		  
		  ' ==============================================
		  ' 4. PARAMETER COMPLEXITY (15%)
		  ' ==============================================
		  Var totalParams As Integer = 0
		  Var totalOptionalParams As Integer = 0
		  Var methodsWithManyParams As Integer = 0
		  
		  For Each method As CodeElement In allMethods
		    totalParams = totalParams + method.ParameterCount
		    totalOptionalParams = totalOptionalParams + method.OptionalParameterCount
		    
		    If method.ParameterCount > 5 Then
		      methodsWithManyParams = methodsWithManyParams + 1
		    End If
		  Next
		  
		  If allMethods.Count > 0 Then
		    score.AverageParameters = totalParams / allMethods.Count
		  Else
		    score.AverageParameters = 0
		  End If
		  
		  ' Score parameters (lower is better)
		  ' Excellent: 0-2 avg = 100
		  ' Good: 2-3 avg = 85
		  ' Fair: 3-4 avg = 70
		  ' Poor: 4-5 avg = 50
		  ' Critical: 5+ avg = 30
		  If score.AverageParameters <= 2 Then
		    score.ParameterScore = 100
		  ElseIf score.AverageParameters <= 3 Then
		    score.ParameterScore = 100 - ((score.AverageParameters - 2) * 15)
		  ElseIf score.AverageParameters <= 4 Then
		    score.ParameterScore = 85 - ((score.AverageParameters - 3) * 15)
		  ElseIf score.AverageParameters <= 5 Then
		    score.ParameterScore = 70 - ((score.AverageParameters - 4) * 20)
		  Else
		    score.ParameterScore = 30
		  End If
		  
		  ' Penalty for methods with too many parameters
		  If allMethods.Count > 0 Then
		    Var excessiveParamPenalty As Double = (methodsWithManyParams / allMethods.Count) * 30
		    score.ParameterScore = score.ParameterScore - excessiveParamPenalty
		    If score.ParameterScore < 0 Then
		      score.ParameterScore = 0
		    End If
		  End If
		  
		  ' ==============================================
		  ' 5. DOCUMENTATION COVERAGE (10%)
		  ' ==============================================
		  Var methodsWithDocs As Integer = 0
		  
		  For Each method As CodeElement In allMethods
		    If method.Code.Trim <> "" Then
		      Var code As String = method.Code
		      
		      ' Look for comment markers at the beginning of the method
		      ' This is a simple heuristic
		      Var lines() As String = code.Split(EndOfLine)
		      Var hasComment As Boolean = False
		      
		      For Each line As String In lines
		        Var trimmed As String = line.Trim
		        
		        ' Check for comment line
		        If trimmed.BeginsWith("'") Or trimmed.BeginsWith("//") Then
		          hasComment = True
		          Exit For
		        End If
		        
		        ' If we hit actual code, stop looking
		        If trimmed <> "" And Not trimmed.BeginsWith("'") And Not trimmed.BeginsWith("//") Then
		          Exit For
		        End If
		      Next
		      
		      If hasComment Then
		        methodsWithDocs = methodsWithDocs + 1
		      End If
		    End If
		  Next
		  
		  If allMethods.Count > 0 Then
		    score.DocumentationCoverage = (methodsWithDocs / allMethods.Count) * 100
		    score.DocumentationScore = score.DocumentationCoverage
		  Else
		    score.DocumentationCoverage = 0
		    score.DocumentationScore = 0
		  End If
		  
		  ' ==============================================
		  ' CALCULATE OVERALL SCORE (WEIGHTED)
		  ' ==============================================
		  score.OverallScore = _
		  (score.ErrorHandlingScore * 0.30) + _
		  (score.ComplexityScore * 0.25) + _
		  (score.CodeReuseScore * 0.20) + _
		  (score.ParameterScore * 0.15) + _
		  (score.DocumentationScore * 0.10)
		  
		  ' Assign grade
		  If score.OverallScore >= 90 Then
		    score.Grade = "A+"
		  ElseIf score.OverallScore >= 85 Then
		    score.Grade = "A"
		  ElseIf score.OverallScore >= 80 Then
		    score.Grade = "A-"
		  ElseIf score.OverallScore >= 75 Then
		    score.Grade = "B+"
		  ElseIf score.OverallScore >= 70 Then
		    score.Grade = "B"
		  ElseIf score.OverallScore >= 65 Then
		    score.Grade = "B-"
		  ElseIf score.OverallScore >= 60 Then
		    score.Grade = "C+"
		  ElseIf score.OverallScore >= 55 Then
		    score.Grade = "C"
		  ElseIf score.OverallScore >= 50 Then
		    score.Grade = "C-"
		  ElseIf score.OverallScore >= 45 Then
		    score.Grade = "D+"
		  ElseIf score.OverallScore >= 40 Then
		    score.Grade = "D"
		  Else
		    score.Grade = "F"
		  End If
		  
		  Return score
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CheckForDuplicates() As String
		  // Find elements with the same name but different full paths
		  Var nameCount As New Dictionary
		  Var report As String = ""
		  
		  // Count occurrences of each name
		  For Each element As CodeElement In AllElements
		    Var name As String = element.Name
		    If nameCount.HasKey(name) Then
		      Var count As Integer = nameCount.Value(name)
		      nameCount.Value(name) = count + 1
		    Else
		      nameCount.Value(name) = 1
		    End If
		  Next
		  
		  // Report duplicates
		  For Each element As CodeElement In AllElements
		    Var name As String = element.Name
		    Var count As Integer = nameCount.Value(name)
		    If count > 1 Then
		      report = report + "Duplicate: " + name + " -> FullPath: " + element.FullPath + EndOfLine
		    End If
		  Next
		  
		  Return report
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CleanCodeForAnalysis(content As String) As String
		  // Remove comments and strings to avoid false positives in code analysis
		  
		  // Remove single-line comments
		  Var rx1 As New RegEx
		  rx1.SearchPattern = "'[^\r\n]*"
		  rx1.ReplacementPattern = ""
		  Var cleanedText As String = rx1.Replace(content)
		  
		  // Remove string literals
		  Var rx2 As New RegEx
		  rx2.SearchPattern = Chr(34) + ".*?" + Chr(34)
		  rx2.ReplacementPattern = ""
		  cleanedText = rx2.Replace(cleanedText)
		  
		  Return cleanedText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CollectCallChain(element As CodeElement, ByRef chain() As CodeElement, visited As Dictionary, currentDepth As Integer, maxDepth As Integer)
		  If element = Nil Or currentDepth > maxDepth Then Return
		  
		  Var key As String = element.FullPath
		  If visited.HasKey(key) Then Return
		  
		  visited.Value(key) = True
		  chain.Add(element)
		  
		  For Each calledElement As CodeElement In element.CallsTo
		    CollectCallChain(calledElement, chain, visited, currentDepth + 1, maxDepth)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  ElementLookup = New Dictionary
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CountOccurrencesInString(text As String, searchFor As String) As Integer
		  // Count occurrences of a substring in text
		  
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
		Private Function CountParametersInList(paramList As String) As Integer
		  // Count parameters in a parameter list, handling nested parentheses
		  
		  If paramList.Trim = "" Then Return 0
		  
		  Var count As Integer = 1  // At least one parameter if string is not empty
		  Var parenDepth As Integer = 0
		  
		  For i As Integer = 0 To paramList.Length - 1
		    Var c As String = paramList.Mid(i, 1)
		    
		    If c = "(" Then
		      parenDepth = parenDepth + 1
		    ElseIf c = ")" Then
		      parenDepth = parenDepth - 1
		    ElseIf c = "," And parenDepth = 0 Then
		      count = count + 1
		    End If
		  Next
		  
		  Return count
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DetectCodeSmells() As CodeSmell()
		  
		  // Function DetectCodeSmells() As CodeSmell()
		  ' Main method to detect all code smells
		  
		  Var smells() As CodeSmell
		  
		  Logger.Log("=== Starting Code Smell Detection ===")
		  
		  ' Detect each type of smell and add to array
		  Var temp() As CodeSmell
		  
		  temp = DetectGodClasses()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectFeatureEnvy()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectMagicNumbers()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectLongParameterLists()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectDeepNesting()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectDeadCode()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  temp = DetectShotgunSurgery()
		  For Each smell As CodeSmell In temp
		    smells.Add(smell)
		  Next
		  
		  Logger.Log("Total code smells detected: " + smells.Count.ToString)
		  
		  DetectedSmells = smells
		  Return smells
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectDatabaseOperations(code As String) As Boolean
		  // Detect database-related operations
		  
		  Var upperCode As String = code.Uppercase
		  
		  // Common Xojo database patterns
		  If upperCode.Contains("SQLSELECT") Then Return True
		  If upperCode.Contains("SQLEXECUTE") Then Return True
		  If upperCode.Contains(".SELECT(") Then Return True
		  If upperCode.Contains(".SELECTSQL") Then Return True
		  If upperCode.Contains(".EXECUTESQL") Then Return True
		  If upperCode.Contains("DATABASE.") Then Return True
		  If upperCode.Contains("RECORDSET") Then Return True
		  If upperCode.Contains(".PREPARE") Then Return True
		  If upperCode.Contains("PREPAREDSTATEMENT") Then Return True
		  If upperCode.Contains(".BIND(") Then Return True
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DetectDeadCode() As CodeSmell()
		  // Private Function DetectDeadCode() As CodeSmell()
		  Logger.Log("=== DETECTDEADCODE STARTED ===")
		  
		  Var allElements() As CodeElement = GetAllElements()
		  Var totalElements As String = allElements.Count.ToString
		  Logger.Log("Total elements to check: " + totalElements)
		  
		  Var codeSmells() As CodeSmell
		  
		  // Collect all code from all elements
		  Var allCode As String = ""
		  For Each element As CodeElement In allElements
		    Var codeCheck As String = element.Code.Trim
		    If codeCheck <> "" Then
		      allCode = allCode + element.Code + EndOfLine
		    End If
		  Next
		  
		  // Clean the code for analysis (remove comments and strings)
		  Var cleanedText As String = CleanCodeForAnalysis(allCode)
		  Var cleanedLength As String = cleanedText.Length.ToString
		  Logger.Log("Cleaned code length: " + cleanedLength + " characters")
		  
		  // Mark which elements are actually used
		  Logger.Log("About to call MarkUsedElements...")
		  MarkUsedElements(cleanedText, allElements)
		  Logger.Log("MarkUsedElements completed")
		  
		  // Count unused elements
		  Var unusedCount As Integer = 0
		  For Each element As CodeElement In allElements
		    If Not element.IsUsed Then
		      unusedCount = unusedCount + 1
		      Var logMsg As String = "UNUSED: " + element.FullPath + " (Type: " + element.ElementType + ")"
		      Logger.Log(logMsg)
		    End If
		  Next
		  
		  Var unusedStr As String = unusedCount.ToString
		  Logger.Log("Total unused elements: " + unusedStr)
		  
		  // Find unused elements and create code smells
		  For Each element As CodeElement In allElements
		    If Not element.IsUsed Then
		      Var smell As New CodeSmell
		      smell.SmellType = "Dead Code"
		      smell.Severity = "Medium"
		      smell.MethodName = element.FullPath
		      Var descPart1 As String = "Unused " + element.ElementType.Lowercase
		      Var descPart2 As String = ": " + element.Name
		      smell.Description = descPart1 + descPart2
		      smell.LineNumber = 0
		      codeSmells.Add(smell)
		    End If
		  Next
		  
		  Logger.Log("=== DETECTDEADCODE COMPLETED ===")
		  Return codeSmells
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectDeepNesting() As CodeSmell()
		  // Private Function DetectDeepNesting() As CodeSmell()
		  ' Detect deeply nested code
		  
		  Var smells() As CodeSmell
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    If method.Code.Trim = "" Then Continue
		    
		    Var maxDepth As Integer = CalculateNestingDepth(method.Code)
		    
		    If maxDepth >= 4 Then
		      Var smell As New CodeSmell
		      smell.SmellType = "Deep Nesting"
		      
		      If maxDepth >= 6 Then
		        smell.Severity = "CRITICAL"
		      ElseIf maxDepth >= 5 Then
		        smell.Severity = "HIGH"
		      Else
		        smell.Severity = "MEDIUM"
		      End If
		      
		      smell.Element = method
		      smell.Description = "Method has " + maxDepth.ToString + " levels of nesting"
		      smell.Details = "Deep nesting makes code hard to understand"
		      smell.Recommendation = "Use guard clauses and extract nested logic into methods"
		      smell.MetricValue = maxDepth
		      smells.Add(smell)
		    End If
		  Next
		  
		  Return smells
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectFeatureEnvy() As CodeSmell()
		  
		  // Private Function DetectFeatureEnvy() As CodeSmell()
		  ' Detect Feature Envy (methods that use other classes more than their own)
		  
		  Var smells() As CodeSmell
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    If method.Code.Trim = "" Then Continue
		    
		    ' Determine the method's class
		    Var methodClass As String = ""
		    Var parts() As String = method.FullPath.Split(".")
		    If parts.Count >= 2 Then
		      methodClass = parts(parts.Count - 2)
		    End If
		    
		    If methodClass = "" Then Continue
		    
		    Var code As String = method.Code
		    Var ownClassReferences As Integer = 0
		    Var otherClassReferences As Integer = 0
		    
		    ' Count "Self." references (own class)
		    ownClassReferences = CountOccurrencesInString(code.Uppercase, "SELF.")
		    ownClassReferences = ownClassReferences + CountOccurrencesInString(code.Uppercase, "ME.")
		    
		    ' Count references to other class properties
		    Var lines() As String = code.Split(EndOfLine)
		    For Each line As String In lines
		      Var trimmed As String = line.Trim.Uppercase
		      
		      ' Skip comments
		      If trimmed.BeginsWith("'") Or trimmed.BeginsWith("//") Then Continue
		      
		      ' Look for dot notation (excluding Self/Me)
		      If trimmed.IndexOf(".") > 0 Then
		        If Not trimmed.Contains("SELF.") And Not trimmed.Contains("ME.") Then
		          ' Count occurrences of dot notation
		          Var dotCount As Integer = 0
		          For i As Integer = 0 To trimmed.Length - 1
		            If trimmed.Mid(i, 1) = "." Then
		              ' Check it's not a numeric decimal
		              If i > 0 And i < trimmed.Length - 1 Then
		                Var before As String = trimmed.Mid(i - 1, 1)
		                Var after As String = trimmed.Mid(i + 1, 1)
		                If Not (before >= "0" And before <= "9" And after >= "0" And after <= "9") Then
		                  dotCount = dotCount + 1
		                End If
		              End If
		            End If
		          Next
		          otherClassReferences = otherClassReferences + dotCount
		        End If
		      End If
		    Next
		    
		    ' Feature envy if uses other classes significantly more than own
		    If otherClassReferences > 10 And otherClassReferences > (ownClassReferences * 3) Then
		      Var smell As New CodeSmell
		      smell.SmellType = "Feature Envy"
		      smell.Severity = If(otherClassReferences > 20, "HIGH", "MEDIUM")
		      smell.Element = method
		      smell.Description = "Method uses other classes' data more than its own"
		      smell.Details = "Own class refs: " + ownClassReferences.ToString + ", Other class refs: " + otherClassReferences.ToString
		      smell.Recommendation = "Move this method to the class it interacts with most, or extract a new class"
		      smell.MetricValue = otherClassReferences
		      smells.Add(smell)
		      
		      Logger.Log("Feature Envy detected: " + method.FullPath)
		    End If
		  Next
		  
		  Return smells
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectFileOperations(code As String) As Boolean
		  // Detect file I/O operations
		  
		  Var upperCode As String = code.Uppercase
		  
		  // File operations
		  If upperCode.Contains("FOLDERITEM") Then Return True
		  If upperCode.Contains("TEXTINPUTSTREAM") Then Return True
		  If upperCode.Contains("TEXTOUTPUTSTREAM") Then Return True
		  If upperCode.Contains("BINARYSTREAM") Then Return True
		  If upperCode.Contains(".READ") And upperCode.Contains("FILE") Then Return True
		  If upperCode.Contains(".WRITE") And upperCode.Contains("FILE") Then Return True
		  If upperCode.Contains(".OPENASTEXT") Then Return True
		  If upperCode.Contains(".CREATEASTEXT") Then Return True
		  If upperCode.Contains(".DELETE") And upperCode.Contains("FILE") Then Return True
		  If upperCode.Contains(".COPY") And upperCode.Contains("FILE") Then Return True
		  If upperCode.Contains(".MOVE") And upperCode.Contains("FILE") Then Return True
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectGodClasses() As CodeSmell()
		  //Private Function DetectGodClasses() As CodeSmell()
		  ' Detect God Classes (classes with too many responsibilities)
		  
		  Var smells() As CodeSmell
		  Var classes() As CodeElement = GetClassElements()
		  
		  For Each classElement As CodeElement In classes
		    Var methodCount As Integer = 0
		    Var totalLOC As Integer = 0
		    
		    ' Count methods in this class
		    Var allMethods() As CodeElement = GetMethodElements()
		    For Each method As CodeElement In allMethods
		      If method.FullPath.BeginsWith(classElement.FullPath + ".") Then
		        methodCount = methodCount + 1
		        totalLOC = totalLOC + method.LinesOfCode
		      End If
		    Next
		    
		    ' Check thresholds
		    Var isGodClass As Boolean = False
		    Var reason As String = ""
		    Var severity As String = ""
		    
		    If methodCount > 25 Then
		      isGodClass = True
		      reason = methodCount.ToString + " methods"
		      severity = "CRITICAL"
		    ElseIf methodCount > 15 Then
		      isGodClass = True
		      reason = methodCount.ToString + " methods"
		      severity = "HIGH"
		    ElseIf totalLOC > 1000 Then
		      isGodClass = True
		      reason = totalLOC.ToString + " lines of code"
		      severity = "CRITICAL"
		    ElseIf totalLOC > 500 Then
		      isGodClass = True
		      reason = totalLOC.ToString + " lines of code"
		      severity = "HIGH"
		    End If
		    
		    If isGodClass Then
		      Var smell As New CodeSmell
		      smell.SmellType = "God Class"
		      smell.Severity = severity
		      smell.Element = classElement
		      smell.Description = "Class has too many responsibilities (" + reason + ")"
		      smell.Details = "Methods: " + methodCount.ToString + ", LOC: " + totalLOC.ToString
		      smell.Recommendation = "Split into smaller, focused classes using Single Responsibility Principle"
		      smell.MetricValue = methodCount
		      smells.Add(smell)
		      
		      Logger.Log("God Class detected: " + classElement.FullPath + " (" + reason + ")")
		    End If
		  Next
		  
		  Return smells
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectLongParameterLists() As CodeSmell()
		  
		  // Private Function DetectLongParameterLists() As CodeSmell()
		  ' Detect methods with too many parameters
		  
		  Var smells() As CodeSmell
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    If method.ParameterCount > 5 Then
		      Var smell As New CodeSmell
		      smell.SmellType = "Long Parameter List"
		      
		      If method.ParameterCount >= 7 Then
		        smell.Severity = "HIGH"
		      Else
		        smell.Severity = "MEDIUM"
		      End If
		      
		      smell.Element = method
		      smell.Description = "Method has " + method.ParameterCount.ToString + " parameters"
		      smell.Details = "Parameters make methods hard to call and maintain"
		      smell.Recommendation = "Use parameter objects or builder pattern"
		      smell.MetricValue = method.ParameterCount
		      smells.Add(smell)
		    End If
		  Next
		  
		  Return smells
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectMagicNumbers() As CodeSmell()
		  
		  // Private Function DetectMagicNumbers() As CodeSmell()
		  ' Detect Magic Numbers (hardcoded numeric constants)
		  
		  Var smells() As CodeSmell
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    If method.Code.Trim = "" Then Continue
		    
		    Var code As String = method.Code
		    Var magicNumbers() As String
		    
		    ' Look for numeric literals in code
		    Var lines() As String = code.Split(EndOfLine)
		    For Each line As String In lines
		      Var trimmed As String = line.Trim
		      
		      ' Skip comments
		      If trimmed.BeginsWith("'") Or trimmed.BeginsWith("//") Then Continue
		      
		      ' Look for numbers in common patterns
		      If trimmed.Contains(" = ") Or trimmed.Contains(" < ") Or trimmed.Contains(" > ") Or _
		        trimmed.Contains(" + ") Or trimmed.Contains(" - ") Or trimmed.Contains(" * ") Or _
		        trimmed.Contains(" / ") Or trimmed.Contains("(") Or trimmed.Contains(",") Then
		        
		        ' Extract potential numbers
		        Var words() As String = trimmed.ReplaceAll(",", " ").ReplaceAll("(", " ").ReplaceAll(")", " ").Split(" ")
		        For Each word As String In words
		          word = word.Trim
		          
		          ' Check if it's a number
		          If word <> "" And IsNumeric(word) Then
		            Var num As Double
		            Try
		              num = Val(word)
		              
		              ' Skip common values
		              If num <> 0 And num <> 1 And num <> -1 And num <> 100 And _
		                num <> 2 And num <> 10 Then
		                ' Check if already in list
		                Var alreadyHave As Boolean = False
		                For Each existing As String In magicNumbers
		                  If existing = word Then
		                    alreadyHave = True
		                    Exit For
		                  End If
		                Next
		                
		                If Not alreadyHave Then
		                  magicNumbers.Add(word)
		                End If
		              End If
		            Catch
		              ' Not a valid number
		            End Try
		          End If
		        Next
		      End If
		    Next
		    
		    ' If significant magic numbers found, report it
		    If magicNumbers.Count >= 3 Then
		      Var smell As New CodeSmell
		      smell.SmellType = "Magic Numbers"
		      smell.Severity = If(magicNumbers.Count >= 5, "MEDIUM", "LOW")
		      smell.Element = method
		      smell.Description = "Method contains " + magicNumbers.Count.ToString + " magic numbers"
		      smell.Details = "Numbers: " + String.FromArray(magicNumbers, ", ")
		      smell.Recommendation = "Replace magic numbers with named constants"
		      smell.MetricValue = magicNumbers.Count
		      smells.Add(smell)
		      
		      Logger.Log("Magic Numbers detected in: " + method.FullPath)
		    End If
		  Next
		  
		  Return smells
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DetectMethodCalls(code As String, callingMethod As CodeElement)
		  // Public Sub DetectMethodCalls(code As String, callingMethod As CodeElement)
		  Logger.Log("=== DetectMethodCalls DEBUG ===")
		  Logger.Log("Caller: " + callingMethod.Name)
		  Logger.Log("Code length: " + code.Length.ToString)
		  
		  If code.Trim = "" Then
		    Logger.Log("  ⚠️ Code is EMPTY!")
		    Return
		  End If
		  
		  // Show first 200 chars of code
		  Var preview As String = code.Left(Min(200, code.Length))
		  Logger.Log("  Preview: " + preview)
		  
		  // Get all methods to check against
		  Var methods() As CodeElement = GetMethodElements()
		  
		  // Check all known methods to see if they're called in this code
		  For Each methodElement As CodeElement In methods
		    Var methodName As String = methodElement.Name
		    Var foundCall As Boolean = False
		    
		    // Pattern 1: MethodName( - with parentheses
		    If code.IndexOf(methodName + "(", ComparisonOptions.CaseInsensitive) >= 0 Then
		      foundCall = True
		    End If
		    
		    // Pattern 2: MethodName followed by whitespace or newline (no parentheses)
		    // But NOT when it's part of a declaration (Sub/Function MethodName)
		    If Not foundCall Then
		      Var patterns() As String = Array( _
		      methodName + " ", _
		      methodName + Chr(13), _
		      methodName + Chr(10), _
		      "." + methodName + " ", _
		      "." + methodName + Chr(13), _
		      "." + methodName + Chr(10) _
		      )
		      
		      For Each pattern As String In patterns
		        If code.IndexOf(pattern, ComparisonOptions.CaseInsensitive) >= 0 Then
		          // Make sure it's not a declaration
		          Var pos As Integer = code.IndexOf(pattern, ComparisonOptions.CaseInsensitive)
		          Var beforeCall As String = code.Left(pos).Trim
		          
		          // Check last 20 chars before the call
		          Var checkLength As Integer = Min(20, beforeCall.Length)
		          Var contextBefore As String = beforeCall.Right(checkLength)
		          
		          // Skip if it looks like a declaration
		          If contextBefore.IndexOf("Sub ", ComparisonOptions.CaseInsensitive) < 0 And _
		            contextBefore.IndexOf("Function ", ComparisonOptions.CaseInsensitive) < 0 Then
		            foundCall = True
		            Exit For
		          End If
		        End If
		      Next
		    End If
		    
		    If foundCall Then
		      // Create relationship if not already exists
		      Var alreadyLinked As Boolean = False
		      For Each existing As CodeElement In callingMethod.CallsTo
		        If existing.FullPath = methodElement.FullPath Then
		          alreadyLinked = True
		          Exit For
		        End If
		      Next
		      
		      If Not alreadyLinked Then
		        callingMethod.CallsTo.Add(methodElement)
		        methodElement.CalledBy.Add(callingMethod)
		        Logger.Log("  Relationship: " + callingMethod.Name + " calls " + methodElement.Name)
		      End If
		    End If
		  Next
		  
		  Logger.Log("  Detected " + callingMethod.CallsTo.Count.ToString + " calls")
		  If callingMethod.CallsTo.Count > 0 Then
		    For Each called As CodeElement In callingMethod.CallsTo
		      Logger.Log("    → " + called.Name)
		    Next
		  End If
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectNetworkOperations(code As String) As Boolean
		  // Detect network-related operations
		  
		  Var upperCode As String = code.Uppercase
		  
		  // Network operations
		  If upperCode.Contains("URLCONNECTION") Then Return True
		  If upperCode.Contains("SOCKET") Then Return True
		  If upperCode.Contains("HTTPSOCKET") Then Return True
		  If upperCode.Contains("TCPSOCKET") Then Return True
		  If upperCode.Contains(".SEND") And upperCode.Contains("HTTP") Then Return True
		  If upperCode.Contains(".GET") And upperCode.Contains("HTTP") Then Return True
		  If upperCode.Contains(".POST") And upperCode.Contains("HTTP") Then Return True
		  If upperCode.Contains("XMLHTTPREQUEST") Then Return True
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectShotgunSurgery() As CodeSmell()
		  // Private Function DetectShotgunSurgery() As CodeSmell()
		  ' Detect Shotgun Surgery (change in one method requires changes in many others)
		  
		  Var smells() As CodeSmell
		  Var methods() As CodeElement = GetMethodElements()
		  
		  For Each method As CodeElement In methods
		    ' Count how many different classes call this method
		    Var callingClasses() As String
		    
		    If method.CalledBy.Count > 10 Then
		      For Each caller As CodeElement In method.CalledBy
		        ' Extract class name from caller path
		        Var parts() As String = caller.FullPath.Split(".")
		        If parts.Count >= 2 Then
		          Var className As String = parts(parts.Count - 2)
		          
		          ' Check if already in list
		          Var alreadyHave As Boolean = False
		          For Each existing As String In callingClasses
		            If existing = className Then
		              alreadyHave = True
		              Exit For
		            End If
		          Next
		          
		          If Not alreadyHave Then
		            callingClasses.Add(className)
		          End If
		        End If
		      Next
		      
		      ' If called by many different classes, it's a shotgun surgery risk
		      If callingClasses.Count >= 5 Then
		        Var smell As New CodeSmell
		        smell.SmellType = "Shotgun Surgery Risk"
		        smell.Severity = If(callingClasses.Count >= 8, "HIGH", "MEDIUM")
		        smell.Element = method
		        smell.Description = "Called by " + method.CalledBy.Count.ToString + " methods across " + callingClasses.Count.ToString + " classes"
		        smell.Details = "Changes here require checking " + callingClasses.Count.ToString + " different classes"
		        smell.Recommendation = "Consider creating a facade or service class to centralize these calls"
		        smell.MetricValue = callingClasses.Count
		        smells.Add(smell)
		        
		        Logger.Log("Shotgun Surgery risk: " + method.FullPath)
		      End If
		    End If
		  Next
		  
		  Return smells
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectTryCatchBlocks(code As String) As Boolean
		  // Detect if code contains Try/Catch error handling
		  
		  Var upperCode As String = code.Uppercase
		  
		  If upperCode.Contains("TRY") And upperCode.Contains("CATCH") Then
		    Return True
		  End If
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DetectTypeConversions(code As String) As Boolean
		  // Detect type conversion operations that could fail
		  
		  Var upperCode As String = code.Uppercase
		  
		  // Type conversions
		  If upperCode.Contains("VAL(") Then Return True
		  If upperCode.Contains("CDBL(") Then Return True
		  If upperCode.Contains("CTYPE(") Then Return True
		  If upperCode.Contains(".TOINTEGER") Then Return True
		  If upperCode.Contains(".TODOUBLE") Then Return True
		  If upperCode.Contains(".FROMSTRING") Then Return True
		  If upperCode.Contains("INTEGER.FROMSTRING") Then Return True
		  If upperCode.Contains("DOUBLE.FROMSTRING") Then Return True
		  If upperCode.Contains("PARSEDATE") Then Return True
		  
		  // Division (potential divide by zero)
		  If upperCode.Contains(" / ") Then Return True
		  If upperCode.Contains(" \\ ") Then Return True
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawArrow(g As Graphics, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
		  // Public Sub DrawArrow(g As Graphics, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
		  // Draw line
		  g.DrawLine(x1, y1, x2, y2)
		  
		  // Calculate angle for arrowhead
		  Var dx As Double = x2 - x1
		  Var dy As Double = y2 - y1
		  Var angle As Double = ATan2(dy, dx)
		  
		  // Draw arrowhead (larger for visibility)
		  Var arrowSize As Integer = 12
		  Var arrowAngle As Double = 0.4  // Angle of arrowhead
		  
		  Var ax1 As Integer = x2 - arrowSize * Cos(angle - arrowAngle)
		  Var ay1 As Integer = y2 - arrowSize * Sin(angle - arrowAngle)
		  Var ax2 As Integer = x2 - arrowSize * Cos(angle + arrowAngle)
		  Var ay2 As Integer = y2 - arrowSize * Sin(angle + arrowAngle)
		  
		  g.DrawLine(x2, y2, ax1, ay1)
		  g.DrawLine(x2, y2, ax2, ay2)
		  
		End Sub
	#tag EndMethod


	#tag Method, Flags = &h21
		Private Sub DrawConnections(g As Graphics, deps() As Dependency, nodes() As MethodNode)
		  // Public Sub DrawConnections(g As Graphics, deps() As Dependency, nodes() As MethodNode)
		  If deps.Count = 0 Then Return
		  
		  g.PenSize = 2  // FIXED: was PenWidth
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  
		  For Each dep As Dependency In deps
		    Var fromNode As MethodNode = Nil
		    Var toNode As MethodNode = Nil
		    
		    For Each node As MethodNode In nodes
		      If node.FullPath = dep.FromPath Then fromNode = node
		      If node.FullPath = dep.ToPath Then toNode = node
		    Next
		    
		    If fromNode <> Nil And toNode <> Nil Then
		      DrawArrow(g, fromNode.X, fromNode.Y, toNode.X, toNode.Y)
		    End If
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawLegend(g As Graphics, x As Integer, y As Integer)
		  // Private Sub DrawLegend(g As Graphics, x As Integer, y As Integer)
		  g.FontSize = 12
		  g.Bold = True
		  g.DrawingColor = Color.RGB(44, 62, 80)
		  g.DrawText("Complexity", x, y)
		  
		  g.FontSize = 10
		  g.Bold = False
		  
		  // Green
		  g.DrawingColor = Color.RGB(46, 204, 113)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 20, 20, 20)
		  g.DrawingColor = Color.Black
		  g.DrawText("Low (1-5)", x + 30, y + 35)
		  
		  // Yellow
		  g.DrawingColor = Color.RGB(241, 196, 15)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 50, 20, 20)
		  g.DrawText("Medium (6-10)", x + 30, y + 65)
		  
		  // Red
		  g.DrawingColor = Color.RGB(231, 76, 60)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 80, 20, 20)
		  g.DrawText("High (11+)", x + 30, y + 95)
		  
		End Sub
	#tag EndMethod


	#tag Method, Flags = &h21
		Private Sub DrawMethodNodes(g As Graphics, nodes() As MethodNode)
		  // Public Sub DrawMethodNodes(g As Graphics, nodes() As MethodNode)
		  For Each node As MethodNode In nodes
		    Var nodeWidth As Integer = 100
		    Var nodeHeight As Integer = 60
		    
		    // Background - use regular rectangle instead of rounded
		    g.DrawingColor = node.NodeColor
		    g.FillRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Border
		    g.DrawingColor = Color.RGB(52, 73, 94)
		    g.PenSize = 2
		    g.DrawRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Method name (shortened if needed)
		    g.FontSize = 9
		    g.Bold = True
		    g.DrawingColor = Color.White
		    
		    Var displayName As String = node.Name
		    If displayName.Length > 15 Then
		      displayName = displayName.Left(12) + "..."
		    End If
		    
		    Var textWidth As Integer = g.TextWidth(displayName)
		    g.DrawText(displayName, node.X - textWidth\2, node.Y - 5)
		    
		    // Metrics (smaller text)
		    g.FontSize = 7
		    g.Bold = False
		    Var metrics As String = "C:" + node.Complexity.ToString + " L:" + node.LOC.ToString
		    Var metricsWidth As Integer = g.TextWidth(metrics)
		    g.DrawText(metrics, node.X - metricsWidth\2, node.Y + 10)
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractClassName(line As String) As String
		  Var className As String = line
		  className = className.Replace("Protected Class ", "")
		  className = className.Replace("Public Class ", "")
		  className = className.Replace("Private Class ", "")
		  className = className.Replace("Class ", "")
		  
		  Var spacePos As Integer = className.IndexOf(" ")
		  If spacePos > 0 Then
		    className = className.Left(spacePos)
		  End If
		  
		  Return className.Trim
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractInterfaceName(declaration As String) As String
		  // Private Function ExtractInterfaceName(declaration As String) As String
		  // Extract interface name from "Interface IMyInterface"
		  Var parts() As String = declaration.Split(" ")
		  
		  For i As Integer = 0 To parts.Count - 1
		    If parts(i).Uppercase = "INTERFACE" And i + 1 < parts.Count Then
		      Return parts(i + 1).Trim
		    End If
		  Next
		  
		  Return "UnknownInterface"
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractMethodName(line As String) As String
		  line = line.Replace("Protected ", "")
		  line = line.Replace("Private ", "")
		  line = line.Replace("Public ", "")
		  line = line.Replace("Sub ", "")
		  line = line.Replace("Function ", "")
		  
		  Var parenPos As Integer = line.IndexOf("(")
		  Var spacePos As Integer = line.IndexOf(" ")
		  Var endPos As Integer = line.Length
		  
		  If parenPos > 0 And (spacePos = -1 Or parenPos < spacePos) Then
		    endPos = parenPos
		  ElseIf spacePos > 0 Then
		    endPos = spacePos
		  End If
		  
		  If endPos > 0 Then
		    Return line.Left(endPos).Trim
		  End If
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractModuleName(line As String) As String
		  Var moduleName As String = line
		  moduleName = moduleName.Replace("Protected Module ", "")
		  moduleName = moduleName.Replace("Public Module ", "")
		  moduleName = moduleName.Replace("Private Module ", "")
		  moduleName = moduleName.Replace("Global Module ", "")
		  moduleName = moduleName.Replace("Module ", "")
		  
		  Var spacePos As Integer = moduleName.IndexOf(" ")
		  If spacePos > 0 Then
		    moduleName = moduleName.Left(spacePos)
		  End If
		  
		  Return moduleName.Trim
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ExtractParameterInfo(signature As String, element As CodeElement)
		  // Private Sub ExtractParameterInfo(signature As String, element As CodeElement)
		  // Extract parameter information from method signature
		  // Signature format: "Sub MethodName(param1 As Type, param2 As Type)" or "Function MethodName(...) As ReturnType"
		  
		  // Find the parameter list (between parentheses)
		  Var openParen As Integer = signature.IndexOf("(")
		  Var closeParen As Integer = signature.IndexOf(")")
		  
		  If openParen < 0 Or closeParen < 0 Or closeParen <= openParen Then
		    // No parameters or malformed signature
		    element.ParameterCount = 0
		    element.OptionalParameterCount = 0
		    element.Parameters = ""
		    Return
		  End If
		  
		  // Extract the parameter list
		  Var paramList As String = signature.Mid(openParen + 1, closeParen - openParen - 1).Trim
		  element.Parameters = paramList
		  
		  If paramList = "" Then
		    // Empty parameter list
		    element.ParameterCount = 0
		    element.OptionalParameterCount = 0
		    Return
		  End If
		  
		  // Count parameters (split by comma, but be careful of nested types)
		  Var params() As String = SplitParameters(paramList)
		  element.ParameterCount = params.Count
		  
		  // Count optional parameters
		  Var optionalCount As Integer = 0
		  For Each param As String In params
		    Var upperParam As String = param.Uppercase
		    // Check for "Optional" keyword or default value "="
		    If upperParam.Contains("OPTIONAL") Or param.Contains("=") Then
		      optionalCount = optionalCount + 1
		    End If
		  Next
		  
		  element.OptionalParameterCount = optionalCount
		  element.ParameterCount =  params.Count
		  Logger.Log("Number of parameters = " + element.ParameterCount.ToString)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractPropertyName(line As String) As String
		  line = line.Replace("Protected Property ", "")
		  line = line.Replace("Private Property ", "")
		  line = line.Replace("Public Property ", "")
		  line = line.Replace("Property ", "")
		  
		  Var spacePos As Integer = line.IndexOf(" ")
		  If spacePos > 0 Then
		    Return line.Left(spacePos).Trim
		  End If
		  
		  Return line.Trim
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractVariableName(line As String) As String
		  line = line.Replace("Dim ", "")
		  line = line.Replace("Var ", "")
		  
		  Var spacePos As Integer = line.IndexOf(" ")
		  Var asPos As Integer = line.IndexOf(" As ")
		  Var equalPos As Integer = line.IndexOf("=")
		  Var endPos As Integer = line.Length
		  
		  If spacePos > 0 Then endPos = spacePos
		  If asPos > 0 And asPos < endPos Then endPos = asPos
		  If equalPos > 0 And equalPos < endPos Then endPos = equalPos
		  
		  If endPos > 0 Then
		    Return line.Left(endPos).Trim
		  End If
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub FinalizeMethod(context As ParsingContext)
		  //Private Sub FinalizeMethod(context As ParsingContext)
		  Logger.Log("=== FinalizeMethod Called ===")
		  Logger.Log("  InMethodOrFunction: " + context.InMethodOrFunction.ToString)
		  Logger.Log("  CurrentMethodFullPath: " + context.CurrentMethodFullPath)
		  
		  If context.InMethodOrFunction And context.CurrentMethodFullPath <> "" Then
		    Logger.Log("  Looking for element: " + context.CurrentMethodFullPath)
		    
		    // Find the element in mElements
		    Var element As CodeElement = FindElementByFullPath(context.CurrentMethodFullPath)
		    
		    If element <> Nil Then
		      Logger.Log("  Element FOUND!")
		      
		      // Store the accumulated code
		      element.Code = context.CurrentMethodCode
		      
		      Logger.Log("  Code length: " + context.CurrentMethodCode.Length.ToString)
		      Logger.Log("  Code lines: " + context.CurrentMethodCode.CountFields(EndOfLine).ToString)
		      
		      // Calculate lines of code
		      element.LinesOfCode = element.Code.CountFields(EndOfLine)
		      
		      Logger.Log("  LOC set to: " + element.LinesOfCode.ToString)
		      
		      // NOW calculate complexity (after code is accumulated)
		      element.CyclomaticComplexity = CalculateMethodComplexity(element)
		      
		      Logger.Log("  Complexity calculated: " + element.CyclomaticComplexity.ToString)
		      
		      // Reset context
		      context.InMethodOrFunction = False
		      context.CurrentMethodFullPath = ""
		      context.CurrentMethodCode = ""
		    Else
		      Logger.Log("  ERROR: Element NOT FOUND!")
		    End If
		  Else
		    Logger.Log("  Skipped - not in method or no path")
		  End If
		  
		  Logger.Log("=== End FinalizeMethod ===")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FindElementByFullPath(fullPath As String) As CodeElement
		  If ElementLookup.HasKey(fullPath) Then
		    Return ElementLookup.Value(fullPath)
		  End If
		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GenerateDependencyGraphPNG(savePath As FolderItem)
		  // Public Sub GenerateDependencyGraphPNG(savePath As FolderItem)
		  // Collect methods and their dependencies
		  Var methodNodes() As MethodNode
		  Var dependencies() As Dependency
		  
		  BuildDependencyData(methodNodes, dependencies)
		  
		  // MUCH WIDER canvas for better layout
		  Var canvasWidth As Integer = 1800  // Increased from 2400
		  Var canvasHeight As Integer = 3000  // Taller to accommodate spreading
		  
		  CalculateGraphLayout(methodNodes, dependencies, canvasWidth, canvasHeight)
		  
		  // Find actual bounds needed
		  Var minX As Integer = 999999
		  Var maxX As Integer = 0
		  Var maxY As Integer = 0
		  
		  For Each node As MethodNode In methodNodes
		    minX = Min(minX, node.X - 60)
		    maxX = Max(maxX, node.X + 60)
		    maxY = Max(maxY, node.Y + 60)
		  Next
		  
		  // Add generous margins
		  Var finalWidth As Integer = maxX - minX + 400
		  Var finalHeight As Integer = maxY + 200
		  
		  // Adjust all nodes to proper position with margins
		  For Each node As MethodNode In methodNodes
		    node.X = node.X - minX + 100
		    node.Y = node.Y + 80
		  Next
		  
		  // Create picture
		  Var pic As New Picture(finalWidth, finalHeight)
		  Var g As Graphics = pic.Graphics
		  
		  // White background
		  g.DrawingColor = Color.White
		  g.FillRectangle(0, 0, finalWidth, finalHeight)
		  
		  // Title
		  g.FontSize = 28  // Bigger title
		  g.Bold = True
		  g.DrawingColor = Color.RGB(44, 62, 80)
		  g.DrawText("Method Dependency Graph", 50, 50)
		  
		  // Draw legend
		  DrawLegend(g, finalWidth - 220, 100)
		  
		  // Draw connections first (behind nodes)
		  g.DrawingColor = Color.RGB(150, 150, 150)  // Lighter gray
		  g.PenSize = 1  // Thinner lines
		  
		  For Each dep As Dependency In dependencies
		    Var fromNode As MethodNode = Nil
		    Var toNode As MethodNode = Nil
		    
		    For Each node As MethodNode In methodNodes
		      If node.FullPath = dep.FromPath Then fromNode = node
		      If node.FullPath = dep.ToPath Then toNode = node
		    Next
		    
		    If fromNode <> Nil And toNode <> Nil Then
		      DrawArrow(g, fromNode.X, fromNode.Y, toNode.X, toNode.Y)
		    End If
		  Next
		  
		  // Draw method nodes - LARGER BOXES
		  For Each node As MethodNode In methodNodes
		    Var nodeWidth As Integer = 140  // Wider (was 100)
		    Var nodeHeight As Integer = 70  // Taller (was 60)
		    
		    // Background
		    g.DrawingColor = node.NodeColor
		    g.FillRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Border
		    g.DrawingColor = Color.RGB(52, 73, 94)
		    g.PenSize = 2
		    g.DrawRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Method name - LARGER FONT
		    g.FontSize = 11  // Bigger (was 9)
		    g.Bold = True
		    g.DrawingColor = Color.White
		    
		    Var displayName As String = node.Name
		    If displayName.Length > 18 Then  // Allow longer names
		      displayName = displayName.Left(15) + "..."
		    End If
		    
		    Var textWidth As Integer = g.TextWidth(displayName)
		    g.DrawText(displayName, node.X - textWidth\2, node.Y - 8)
		    
		    // Metrics - LARGER FONT
		    g.FontSize = 9  // Bigger (was 7)
		    g.Bold = False
		    Var metrics As String = "C:" + node.Complexity.ToString + " L:" + node.LOC.ToString
		    Var metricsWidth As Integer = g.TextWidth(metrics)
		    g.DrawText(metrics, node.X - metricsWidth\2, node.Y + 15)
		  Next
		  
		  // Save as PNG
		  Var outputFile As FolderItem = savePath.Child("dependency_graph.png")
		  pic.Save(outputFile, Picture.Formats.PNG)
		  outputFile.Launch()
		  
		  MessageBox("Dependency graph saved as PNG!")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GenerateEdgesJSON(deps() As Dependency) As String
		  // Private Function GenerateEdgesJSON(deps() As Dependency) As String
		  Var json As String = "["
		  Var edgeCount As Integer = 0
		  
		  // Build a lookup of fullPath to first matching index in methodNodes
		  Var pathToIndex As New Dictionary
		  
		  Var nodeIndex As Integer = 0
		  For Each element As CodeElement In mElements
		    If element.ElementType = "METHOD" Then
		      Var key As String = element.FullPath
		      // Store only first occurrence of each path
		      If Not pathToIndex.HasKey(key) Then
		        pathToIndex.Value(key) = nodeIndex
		      End If
		      nodeIndex = nodeIndex + 1
		    End If
		  Next
		  
		  // Generate edges
		  For i As Integer = 0 To deps.Count - 1
		    Var dep As Dependency = deps(i)
		    
		    // Find indices for from and to
		    If pathToIndex.HasKey(dep.FromPath) And pathToIndex.HasKey(dep.ToPath) Then
		      Var fromIdx As Integer = pathToIndex.Value(dep.FromPath)
		      Var toIdx As Integer = pathToIndex.Value(dep.ToPath)
		      
		      Var safeFrom As String = dep.FromPath.ReplaceAll("""", "\""")
		      Var safeTo As String = dep.ToPath.ReplaceAll("""", "\""")
		      
		      If edgeCount > 0 Then json = json + ","
		      
		      json = json + "{"
		      json = json + """from"": """ + safeFrom + "_" + fromIdx.ToString + ""","
		      json = json + """to"": """ + safeTo + "_" + toIdx.ToString + ""","
		      json = json + """arrows"": ""to"","
		      json = json + """color"": {""color"": ""#95a5a6"", ""highlight"": ""#3498db""},"
		      json = json + """smooth"": {""type"": ""cubicBezier""}"
		      json = json + "}"
		      
		      edgeCount = edgeCount + 1
		    End If
		  Next
		  
		  json = json + "]"
		  Return json
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GenerateInteractiveDependencyGraph(savePath As FolderItem)
		  // Public Sub GenerateInteractiveDependencyGraph(savePath As FolderItem)
		  // Collect data
		  Var methodNodes() As MethodNode
		  Var dependencies() As Dependency
		  
		  BuildDependencyData(methodNodes, dependencies)
		  
		  // Generate JSON data
		  Var nodesJSON As String = GenerateNodesJSON(methodNodes)
		  Var edgesJSON As String = GenerateEdgesJSON(dependencies)
		  
		  // Build HTML with embedded Vis.js
		  Var html As String = BuildInteractiveHTML(nodesJSON, edgesJSON)
		  
		  // Save and open
		  Var outputFile As FolderItem = savePath.Child("dependency_graph.html")
		  Var stream As TextOutputStream
		  
		  Try
		    stream = TextOutputStream.Create(outputFile)
		    stream.Write(html)
		    stream.Close()
		    outputFile.Launch()
		    MessageBox("Interactive dependency graph created!")
		  Catch e As IOException
		    MessageBox("Error saving file: " + e.Message)
		  End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GenerateNodesJSON(nodes() As MethodNode) As String
		  // Private Function GenerateNodesJSON(nodes() As MethodNode) As String
		  Var json As String = "["
		  
		  For i As Integer = 0 To nodes.Count - 1
		    Var node As MethodNode = nodes(i)
		    
		    Var safeName As String = node.Name.ReplaceAll("""", "\""")
		    Var safePath As String = node.FullPath.ReplaceAll("""", "\""")
		    Var uniqueId As String = safePath + "_" + i.ToString
		    
		    // Color by complexity
		    Var colorHex As String
		    If node.Complexity <= 5 Then
		      colorHex = "#2ecc71"
		    ElseIf node.Complexity <= 10 Then
		      colorHex = "#f1c40f"
		    Else
		      colorHex = "#e74c3c"
		    End If
		    
		    // SMALLER size based on LOC
		    Var nodeSize As Integer = 10 + Min(node.LOC, 50) \ 5
		    
		    json = json + "{"
		    json = json + """id"": """ + uniqueId + ""","
		    json = json + """label"": """ + safeName + ""","
		    json = json + """title"": ""<b>" + safeName + "</b><br>Complexity: " + node.Complexity.ToString + "<br>LOC: " + node.LOC.ToString + "<br>Path: " + safePath + ""","
		    json = json + """color"": {""background"": """ + colorHex + """, ""border"": ""#2c3e50"", ""highlight"": {""background"": """ + colorHex + """, ""border"": ""#000000""}},"
		    json = json + """size"": " + nodeSize.ToString + ","
		    json = json + """font"": {""color"": ""white"", ""size"": 10, ""face"": ""Arial""},"  // Smaller font
		    json = json + """borderWidth"": 2,"
		    json = json + """complexity"": " + node.Complexity.ToString + ","
		    json = json + """loc"": " + node.LOC.ToString + ","
		    json = json + """fullPath"": """ + safePath + """"
		    json = json + "}"
		    
		    If i < nodes.Count - 1 Then
		      json = json + ","
		    End If
		  Next
		  
		  json = json + "]"
		  Return json
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetAllElements() As CodeElement()
		  // Public Function GetAllElements() As CodeElement()
		  Return mElements
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetAverageParametersPerMethod() As Double
		  //Public Function GetAverageParametersPerMethod() As Double
		  Var methods() As CodeElement = GetMethodElements()
		  
		  If methods.Count = 0 Then
		    Return 0.0
		  End If
		  
		  Var totalParams As Integer = 0
		  For Each method As CodeElement In methods
		    totalParams = totalParams + method.ParameterCount
		  Next
		  
		  Return totalParams / methods.Count
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCallChain(startElement As CodeElement, maxDepth As Integer = 5) As CodeElement()
		  Var chain() As CodeElement
		  Var visited As New Dictionary
		  
		  CollectCallChain(startElement, chain, visited, 0, maxDepth)
		  
		  Return chain
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetClassElements() As CodeElement()
		  // Public Function GetClassElements() As CodeElement()
		  Var classes() As CodeElement
		  
		  For Each element As CodeElement In mElements
		    If element.ElementType = "CLASS" Then
		      classes.Add(element)
		    End If
		  Next
		  
		  Return classes
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetErrorHandlingStats() As Dictionary
		  // GetErrorHandlingStats() As Dicgtionary
		  // Calculate error handling statistics
		  
		  Var stats As New Dictionary
		  
		  Var totalMethods As Integer = 0
		  Var methodsWithTryCatch As Integer = 0
		  Var methodsWithRiskyPatterns As Integer = 0
		  Var highRiskCount As Integer = 0
		  Var mediumRiskCount As Integer = 0
		  
		  Var riskyMethods() As CodeElement
		  
		  For Each element As CodeElement In GetMethodElements()  
		    If element.Code.Trim = "" Then Continue
		    
		    totalMethods = totalMethods + 1
		    
		    If element.HasTryCatch Then
		      methodsWithTryCatch = methodsWithTryCatch + 1
		    End If
		    
		    If element.RiskyPatterns.Count > 0 Then
		      methodsWithRiskyPatterns = methodsWithRiskyPatterns + 1
		      riskyMethods.Add(element)
		      
		      For Each pattern As ErrorPattern In element.RiskyPatterns
		        If pattern.RiskLevel = "HIGH" Then
		          highRiskCount = highRiskCount + 1
		        ElseIf pattern.RiskLevel = "MEDIUM" Then
		          mediumRiskCount = mediumRiskCount + 1
		        End If
		      Next
		    End If
		  Next
		  
		  Var coveragePercent As Integer = 0
		  If totalMethods > 0 Then
		    coveragePercent = (methodsWithTryCatch * 100) \ totalMethods
		  End If
		  
		  stats.Value("totalMethods") = totalMethods
		  stats.Value("methodsWithTryCatch") = methodsWithTryCatch
		  stats.Value("methodsWithRiskyPatterns") = methodsWithRiskyPatterns
		  stats.Value("coveragePercent") = coveragePercent
		  stats.Value("highRiskCount") = highRiskCount
		  stats.Value("mediumRiskCount") = mediumRiskCount
		  stats.Value("riskyMethods") = riskyMethods
		  
		  Return stats
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetMaxParameters() As Integer
		  // Public Function GetMaxParameters() As Integer
		  Var methods() As CodeElement = GetMethodElements()
		  Var maxParams As Integer = 0
		  
		  For Each method As CodeElement In methods
		    If method.ParameterCount > maxParams Then
		      maxParams = method.ParameterCount
		    End If
		  Next
		  
		  Return maxParams
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetMethodElements() As CodeElement()
		  // Public Function GetMethodElements() As CodeElement()
		  Var methods() As CodeElement
		  
		  For Each element As CodeElement In mElements
				    
		    If element.ElementType = "METHOD" Then  // ← Is this matching?
		      methods.Add(element)
		    End If
		  Next
		  
		  Logger.Log("Found " + methods.Count.ToString + " methods")
		  Return methods
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetMethodsExceedingParameterThreshold(threshold As Integer) As Integer
		  Var methods() As CodeElement = GetMethodElements()
		  Var count As Integer = 0
		  
		  For Each method As CodeElement In methods
		    If method.ParameterCount > threshold Then
		      count = count + 1
		    End If
		  Next
		  
		  Return count
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetMethodsWithHighParameterCount(threshold As Integer) As CodeElement()
		  // Public Function GetMethodsWithHighParameterCount(threshold As Integer) As CodeElement()
		  Var methods() As CodeElement = GetMethodElements()
		  Var highParamMethods() As CodeElement
		  
		  For Each method As CodeElement In methods
		    If method.ParameterCount > threshold Then
		      highParamMethods.Add(method)
		    End If
		  Next
		  
		  Return highParamMethods
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetModuleElements() As CodeElement()
		  // Public Function GetModuleElements() As CodeElement()
		  Var modules() As CodeElement
		  
		  For Each element As CodeElement In mElements
		    If element.ElementType = "MODULE" Then
		      modules.Add(element)
		    End If
		  Next
		  
		  Return modules
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetTotalOptionalParameters() As Integer
		  Var methods() As CodeElement = GetMethodElements()
		  Var totalOptional As Integer = 0
		  
		  For Each method As CodeElement In methods
		    totalOptional = totalOptional + method.OptionalParameterCount
		  Next
		  
		  Return totalOptional
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetUnusedElements() As CodeElement()
		  // Public Function GetUnusedElements() As CodeElement()
		  Var unused() As CodeElement
		  
		  Var allElements() As CodeElement = GetAllElements()
		  
		  For Each element As CodeElement In allElements
		    If Not element.IsUsed Then
		      unused.Add(element)
		    End If
		  Next
		  
		  Return unused
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleEndTag(line As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentInterface As String)
		  // Handle different types of end tags
		  
		  If line.BeginsWith("#tag EndModule") Then
		    currentModule = ""
		    currentClass = ""
		    currentInterface = ""
		    ResetMethodContext(inMethodOrFunction, currentMethodFullPath, currentMethodCode)
		    
		  ElseIf line.BeginsWith("#tag EndClass") Then
		    currentClass = ""
		    currentInterface = ""
		    ResetMethodContext(inMethodOrFunction, currentMethodFullPath, currentMethodCode)
		    
		  ElseIf line.BeginsWith("#tag End") Or line = "End" Or line = "End Sub" Or line = "End Function" Then
		    // Save the accumulated code to the method element
		    If inMethodOrFunction And currentMethodFullPath.Trim <> "" Then
		      Var methodElement As CodeElement = FindElementByFullPath(currentMethodFullPath)
		      If methodElement <> Nil Then
		        methodElement.Code = currentMethodCode
		        Logger.Log("Stored code for method: " + currentMethodFullPath + " (" + currentMethodCode.Length.ToString + " chars)")
		      End If
		    End If
		    
		    ResetMethodContext(inMethodOrFunction, currentMethodFullPath, currentMethodCode)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsClassDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Protected Class ") Or _
		  line.BeginsWith("Public Class ") Or _
		  line.BeginsWith("Private Class ") Or _
		  line.BeginsWith("Class ")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsEndMethodStatement(line As String) As Boolean
		  // Private Function IsEndMethodStatement(line As String) As Boolean
		  // Check if line marks the end of a method/function
		  
		  Var upper As String = line.Uppercase.Trim
		  
		  Return upper = "END" Or _
		  upper = "END SUB" Or _
		  upper = "END FUNCTION" Or _
		  upper = "END GET" Or _
		  upper = "END SET" Or _
		  upper.BeginsWith("END ")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsEndTag(line As String) As Boolean
		  // Check if line is any kind of end tag
		  Return line.BeginsWith("#tag End") Or _
		  line = "End" Or _
		  line = "End Sub" Or _
		  line = "End Function"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsInterfaceDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Protected Interface ") Or _
		  line.BeginsWith("Public Interface ") Or _
		  line.BeginsWith("Private Interface ") Or _
		  line.BeginsWith("Interface ")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsMethodOrFunctionDeclaration(line As String) As Boolean
		  // Private Function IsMethodOrFunctionDeclaration(line As String) As Boolean
		  // Check for proper method/function declarations at the start of the line
		  Return (line.BeginsWith("Sub ") Or _
		  line.BeginsWith("Function ") Or _
		  line.BeginsWith("Protected Sub ") Or _
		  line.BeginsWith("Protected Function ") Or _
		  line.BeginsWith("Private Sub ") Or _
		  line.BeginsWith("Private Function ") Or _
		  line.BeginsWith("Public Sub ") Or _
		  line.BeginsWith("Public Function ") Or _
		  line.BeginsWith("Shared Sub ") Or _
		  line.BeginsWith("Shared Function ")) And _
		  Not line.BeginsWith("#tag")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsModuleDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Protected Module ") Or _
		  line.BeginsWith("Public Module ") Or _
		  line.BeginsWith("Private Module ") Or _
		  line.BeginsWith("Global Module ") Or _
		  line.BeginsWith("Module ")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsNumeric(value As String) As Boolean
		  // Private Function IsNumeric(value As String) As Boolean
		  ' Helper to check if a string is numeric
		  
		  If value.Trim = "" Then Return False
		  
		  Try
		    Var test As Double = Val(value)
		    Return True
		  Catch
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsPropertyDeclaration(line As String) As Boolean
		  // Private Function IsPropertyDeclaration(line As String) As Boolean
		  // Check for proper property declarations at the start of the line
		  Return (line.BeginsWith("Property ") Or _
		  line.BeginsWith("Protected Property ") Or _
		  line.BeginsWith("Private Property ") Or _
		  line.BeginsWith("Public Property ") Or _
		  line.BeginsWith("Shared Property ")) And _
		  Not line.BeginsWith("#tag")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsSystemMethod(methodName As String) As Boolean
		  Var systemMethods() As String = Array("Open", "Close", "Constructor", "Destructor", _
		  "Paint", "MouseDown", "MouseUp", "MouseDrag", "MouseMove", "MouseEnter", "MouseExit", _
		  "KeyDown", "KeyUp", "GotFocus", "LostFocus", "EnableMenuItems", "MenuBarSelected", _
		  "Activate", "Deactivate", "Resized", "Moved", "Opening", "Closing", _
		  "Action", "Pressed", "TextChanged", "SelectionChanged", "DropObject", _
		  "DragEnter", "DragExit", "DragOver", "DragWithin")
		  
		  Return systemMethods.IndexOf(methodName) >= 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsVariableDeclaration(line As String) As Boolean
		  Return (line.BeginsWith("Var ") Or line.BeginsWith("Dim ")) And Not line.BeginsWith("#tag")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsXojoSourceFile(item As FolderItem) As Boolean
		  Var name As String = item.Name.Lowercase
		  Return name.EndsWith(".xojo_code") Or _
		  name.EndsWith(".xojo_window") Or _
		  name.EndsWith(".xojo_menu") Or _
		  name.EndsWith(".xojo_toolbar") Or _
		  name.Contains(".xojo_")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub MarkUsedElements(cleanedText As String, allElements() As CodeElement)
		  // Private Sub MarkUsedElements(cleanedText As String, allElements() As CodeElement)
		  Logger.Log("=== MARKUSEDELEMENTS STARTED ===")
		  Var totalStr As String = allElements.Count.ToString
		  Logger.Log("Total elements to check: " + totalStr)
		  Var cleanedLengthStr As String = cleanedText.Length.ToString
		  Logger.Log("Cleaned text length: " + cleanedLengthStr)
		  
		  // Count system methods (event handlers)
		  Var systemMethodCount As Integer = 0
		  For Each element As CodeElement In allElements
		    If element.ElementType = "METHOD" And IsSystemMethod(element.Name) Then
		      Var countPlus As Integer = systemMethodCount + 1
		      systemMethodCount = countPlus
		      Logger.Log("Event handler: " + element.FullPath)
		    End If
		  Next
		  
		  Var handlerCountStr As String = systemMethodCount.ToString
		  Logger.Log("Total event handlers found: " + handlerCountStr)
		  
		  Var unusedEventHandlers() As String
		  
		  For Each element As CodeElement In allElements
		    Var name As String = element.Name
		    
		    Select Case element.ElementType
		    Case "METHOD"
		      If IsSystemMethod(name) Then
		        // Event handlers: check if explicitly called
		        Var fullName As String = element.FullPath
		        Var searchPattern1 As String = fullName + "("
		        Var searchPattern2 As String = "AddressOf " + fullName
		        
		        If cleanedText.IndexOf(searchPattern1) >= 0 Or cleanedText.IndexOf(searchPattern2) >= 0 Then
		          element.IsUsed = True
		          Logger.Log("Event handler USED: " + fullName)
		        Else
		          element.IsUsed = False
		          unusedEventHandlers.Add(fullName)
		          Logger.Log("Event handler UNUSED: " + fullName)
		        End If
		      Else
		        // Regular methods: check by name
		        Var searchPattern1 As String = name + "("
		        Var searchPattern2 As String = "AddressOf " + name
		        
		        If cleanedText.IndexOf(searchPattern1) >= 0 Or cleanedText.IndexOf(searchPattern2) >= 0 Then
		          element.IsUsed = True
		        Else
		          element.IsUsed = False
		        End If
		      End If
		      
		    Case "CLASS"
		      Var searchPattern1 As String = "New " + name
		      Var searchPattern2 As String = "As " + name
		      Var searchPattern3 As String = "Inherits " + name
		      
		      If cleanedText.IndexOf(searchPattern1) >= 0 Or _
		        cleanedText.IndexOf(searchPattern2) >= 0 Or _
		        cleanedText.IndexOf(searchPattern3) >= 0 Then
		        element.IsUsed = True
		      Else
		        element.IsUsed = False
		      End If
		      
		    Case "MODULE"
		      Var searchPattern As String = name + "."
		      If cleanedText.IndexOf(searchPattern) >= 0 Then
		        element.IsUsed = True
		      Else
		        element.IsUsed = False
		      End If
		      
		    Case "PROPERTY", "VARIABLE"
		      Var searchPattern1 As String = "." + name
		      Var searchPattern2 As String = " " + name + " "
		      Var searchPattern3 As String = name + " ="
		      
		      If cleanedText.IndexOf(searchPattern1) >= 0 Or _
		        cleanedText.IndexOf(searchPattern2) >= 0 Or _
		        cleanedText.IndexOf(searchPattern3) >= 0 Then
		        element.IsUsed = True
		      Else
		        element.IsUsed = False
		      End If
		    End Select
		  Next
		  
		  // Log summary
		  Var unusedCountStr As String = unusedEventHandlers.Count.ToString
		  Logger.Log("Found " + unusedCountStr + " unused event handlers")
		  For Each handler As String In unusedEventHandlers
		    Logger.Log("  - " + handler)
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ParseFileContent(content As String, fileName As String)
		  //Private Sub ParseFileContent(content As String, fileName As String)
		  // Main orchestrator - delegates to focused helper methods
		  
		  Var lines() As String = content.Split(EndOfLine)
		  Var context As New ParsingContext
		  context.FileName = fileName
		  
		  For Each line As String In lines
		    Var trimmedLine As String = line.Trim
		    
		    // Skip empty lines (but keep them in method code if we're in a method)
		    If trimmedLine = "" Then
		      AccumulateCodeLine(context, line)
		      Continue
		    End If
		    
		    // Process the line based on what it declares
		    If IsModuleDeclaration(trimmedLine) Then
		      ProcessModuleDeclaration(trimmedLine, context)
		      
		    ElseIf IsInterfaceDeclaration(trimmedLine) Then
		      ProcessInterfaceDeclaration(trimmedLine, context)
		      
		    ElseIf IsClassDeclaration(trimmedLine) Then
		      ProcessClassDeclaration(trimmedLine, context)
		      
		    ElseIf IsMethodOrFunctionDeclaration(trimmedLine) Then
		      ProcessMethodDeclaration(trimmedLine, context)
		      
		    ElseIf IsPropertyDeclaration(trimmedLine) Then
		      ProcessPropertyDeclaration(trimmedLine, context)
		      
		    ElseIf IsEndMethodStatement(trimmedLine) Then
		      FinalizeMethod(context)
		      
		    Else
		      // Regular code line - accumulate if we're in a method
		      AccumulateCodeLine(context, line)
		    End If
		  Next
		  
		  // Handle any method still open at end of file
		  If context.InMethodOrFunction Then
		    FinalizeMethod(context)
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParseMethodParameters(methodCode As String) As Dictionary
		  // Parse method signature and extract parameter information
		  // Returns Dictionary with: parameterCount, optionalCount
		  
		  Var result As New Dictionary
		  result.Value("parameterCount") = 0
		  result.Value("optionalCount") = 0
		  
		  If methodCode.Trim = "" Then Return result
		  
		  // Find the method declaration line (first line that's not a comment)
		  Var lines() As String = methodCode.Split(EndOfLine)
		  Var declarationLine As String = ""
		  
		  For Each line As String In lines
		    Var trimmedLine As String = line.Trim
		    If trimmedLine.Length > 0 And Not trimmedLine.BeginsWith("'") And Not trimmedLine.BeginsWith("//") Then
		      declarationLine = trimmedLine
		      Exit For line
		    End If
		  Next
		  
		  If declarationLine = "" Then Return result
		  
		  // Check if this is a method declaration (Sub or Function)
		  Var upperLine As String = declarationLine.Uppercase
		  If Not (upperLine.BeginsWith("SUB ") Or upperLine.BeginsWith("FUNCTION ") Or _
		    upperLine.BeginsWith("PRIVATE SUB ") Or upperLine.BeginsWith("PRIVATE FUNCTION ") Or _
		    upperLine.BeginsWith("PUBLIC SUB ") Or upperLine.BeginsWith("PUBLIC FUNCTION ") Or _
		    upperLine.BeginsWith("PROTECTED SUB ") Or upperLine.BeginsWith("PROTECTED FUNCTION ")) Then
		    Return result
		  End If
		  
		  // Find parameter list (between parentheses)
		  Var openParen As Integer = declarationLine.IndexOf("(")
		  Var closeParen As Integer = declarationLine.IndexOf(")")
		  
		  If openParen < 0 Or closeParen < 0 Or closeParen <= openParen Then
		    Return result  // No parameters or invalid syntax
		  End If
		  
		  Var paramList As String = declarationLine.Mid(openParen + 1, closeParen - openParen - 1).Trim
		  
		  If paramList = "" Then
		    Return result  // Empty parameter list
		  End If
		  
		  // Count parameters by splitting on commas (ignoring commas inside parentheses)
		  Var paramCount As Integer = CountParametersInList(paramList)
		  result.Value("parameterCount") = paramCount
		  
		  // Count optional parameters
		  Var optionalCount As Integer = CountOccurrencesInString(paramList.Uppercase, "OPTIONAL ")
		  result.Value("optionalCount") = optionalCount
		  
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessClassDeclaration(declaration As String, context As ParsingContext)
		  // Private Sub ProcessClassDeclaration(declaration As String, context As ParsingContext)
		  // Handle class declarations
		  
		  context.CurrentClass = ExtractClassName(declaration)
		  context.CurrentInterface = ""
		  
		  Var fullPath As String = BuildFullPath(context.CurrentModule, "", context.CurrentClass)
		  Var element As New CodeElement("CLASS", context.CurrentClass, fullPath, context.FileName)
		  mElements.Add(element)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessInterfaceDeclaration(declaration As String, context As ParsingContext)
		  // Private Sub ProcessInterfaceDeclaration(declaration As String, context As ParsingContext)
		  // Handle interface declarations
		  
		  context.CurrentInterface = ExtractInterfaceName(declaration)
		  context.CurrentClass = context.CurrentInterface
		  
		  Var fullPath As String = BuildFullPath(context.CurrentModule, "", context.CurrentInterface)
		  Var element As New CodeElement("INTERFACE", context.CurrentInterface, fullPath, context.FileName)
		  mElements.Add(element)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessLine(trimmedLine As String, originalLine As String, context As ParsingContext)
		  // Private Sub ProcessLine(trimmedLine As String, originalLine As String, context As ParsingContext)
		  If IsModuleDeclaration(trimmedLine) Then
		    ProcessModuleDeclaration(trimmedLine, context)
		    
		  ElseIf IsInterfaceDeclaration(trimmedLine) Then
		    ProcessInterfaceDeclaration(trimmedLine, context)
		    
		  ElseIf IsClassDeclaration(trimmedLine) Then
		    ProcessClassDeclaration(trimmedLine, context)
		    
		  ElseIf IsMethodOrFunctionDeclaration(trimmedLine) Then
		    ProcessMethodDeclaration(trimmedLine, context)
		    
		  ElseIf IsPropertyDeclaration(trimmedLine) Then
		    ProcessPropertyDeclaration(trimmedLine, context)
		    
		  ElseIf IsEndMethodStatement(trimmedLine) Then
		    FinalizeMethod(context)
		    
		  Else
		    // Regular code line
		    If context.InMethodOrFunction Then
		      context.CurrentMethodCode = context.CurrentMethodCode + originalLine + EndOfLine
		    End If
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessLineForRelationships(line As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentMethodFullPath As String, ByRef inMethod As Boolean)
		  // Private Sub ProcessLineForRelationships(line As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentMethodFullPath As String, ByRef inMethod As Boolean)
		  // Process a line to track context and detect relationships
		  
		  If IsModuleDeclaration(line) Then
		    currentModule = ExtractModuleName(line)
		    currentClass = ""
		    currentMethodFullPath = ""
		    inMethod = False
		    
		  ElseIf IsClassDeclaration(line) Then
		    currentClass = ExtractClassName(line)
		    currentMethodFullPath = ""
		    inMethod = False
		    
		  ElseIf IsMethodOrFunctionDeclaration(line) Then
		    currentMethodFullPath = BuildMethodFullPath(line, currentModule, currentClass)
		    inMethod = True
		  End If
		  
		  // Don't call DetectMethodCalls here!
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessMethodDeclaration(declaration As String, context As ParsingContext)
		  // Private Sub ProcessMethodDeclaration(declaration As String, context As ParsingContext)
		  Logger.Log(">>> ProcessMethodDeclaration CALLED with: " + declaration)
		  
		  // Finalize any previous method that was still open
		  If context.InMethodOrFunction Then
		    FinalizeMethod(context)
		  End If
		  
		  Var methodName As String = ExtractMethodName(declaration)
		  Logger.Log(">>> Method name extracted: " + methodName)
		  
		  Var fullPath As String = BuildFullPath(context.CurrentModule, context.CurrentClass, methodName)
		  Logger.Log(">>> Full path: " + fullPath)
		  
		  Var element As New CodeElement("METHOD", methodName, fullPath, context.FileName)
		  
		  // Extract parameter information from the signature
		  ExtractParameterInfo(declaration, element)
		  
		  // ADD TO BOTH THE ARRAY AND THE DICTIONARY:
		  mElements.Add(element)
		  ElementLookup.Value(fullPath) = element  // ← ADD THIS LINE!
		  
		  // Start tracking this method
		  context.InMethodOrFunction = True
		  context.CurrentMethodFullPath = fullPath
		  context.CurrentMethodCode = ""
		  
		  Logger.Log(">>> Method declaration processed successfully")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessModuleDeclaration(declaration As String, context As ParsingContext)
		  // Private Sub ProcessModuleDeclaration(declaration As String, context As ParsingContext)
		  // Handle module declarations
		  
		  context.CurrentModule = ExtractModuleName(declaration)
		  context.CurrentClass = ""
		  context.CurrentInterface = ""
		  
		  Var fullPath As String = context.CurrentModule
		  Var element As New CodeElement("MODULE", context.CurrentModule, fullPath, context.FileName)
		  mElements.Add(element)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessPropertyDeclaration(declaration As String, context As ParsingContext)
		  // Private Sub ProcessPropertyDeclaration(declaration As String, context As ParsingContext)
		  // Handle property declarations
		  
		  Var propertyName As String = ExtractPropertyName(declaration)
		  Var fullPath As String = BuildFullPath(context.CurrentModule, context.CurrentClass, propertyName)
		  
		  Var element As New CodeElement("PROPERTY", propertyName, fullPath, context.FileName)
		  mElements.Add(element)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessSourceFile(item As FolderItem)
		  Try
		    Var tis As TextInputStream = TextInputStream.Open(item)
		    If tis = Nil Then Return
		    
		    Var content As String = tis.ReadAll
		    tis.Close
		    
		    ParseFileContent(content, item.Name)
		    
		  Catch e As IOException
		    Logger.Log("Error reading file: " + item.Name + " - " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessVariableDeclaration(line As String, currentModule As String, currentClass As String, fileName As String)
		  Var varName As String = ExtractVariableName(line)
		  
		  If varName.Trim <> "" Then
		    Var fullPath As String = ""
		    
		    If currentModule.Trim <> "" And currentClass.Trim <> "" Then
		      fullPath = currentModule + "." + currentClass + "." + varName
		    ElseIf currentModule.Trim <> "" Then
		      fullPath = currentModule + "." + varName
		    ElseIf currentClass.Trim <> "" Then
		      fullPath = currentClass + "." + varName
		    Else
		      fullPath = varName
		    End If
		    
		    If Not ElementLookup.HasKey(fullPath) Then
		      Var element As New CodeElement("VARIABLE", varName, fullPath, fileName)
		      mElements.Add(element)
		      Logger.Log("    Total elements: " + mElements.Count.ToString)
		      ElementLookup.Value(fullPath) = element
		      Logger.Log("Found variable: " + fullPath + " in " + fileName)
		    End If
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ResetMethodContext(ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String)
		  // Reset method tracking variables
		  inMethodOrFunction = False
		  currentMethodFullPath = ""
		  currentMethodCode = ""
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ScanFileForRelationships(item As FolderItem)
		  // Private Sub ScanFileForRelationships(item As FolderItem)
		  // Guard clauses - check if we can actually read this item
		  If item = Nil Then
		    Logger.Log("⚠️ ScanFileForRelationships: item is Nil")
		    Return
		  End If
		  
		  If Not item.Exists Then
		    Logger.Log("⚠️ ScanFileForRelationships: item does not exist")
		    Return
		  End If
		  
		  If item.IsFolder Then
		    Logger.Log("⚠️ ScanFileForRelationships: item is a folder, not a file: " + item.Name)
		    Return
		  End If
		  
		  // Only process Xojo source files
		  If Not IsXojoSourceFile(item) Then
		    Return
		  End If
		  
		  Try
		    Var tis As TextInputStream = TextInputStream.Open(item)
		    Var content As String = tis.ReadAll
		    tis.Close
		    
		    // Clean the code
		    Var cleanedText As String = CleanCodeForAnalysis(content)
		    
		    // Analyze for relationships
		    AnalyzeFileForRelationships(content)
		    
		    // Get all elements and mark them as used
		    Var allElements() As CodeElement = GetAllElements()
		    MarkUsedElements(cleanedText, allElements)
		    
		  Catch e As IOException
		    Logger.Log("❌ Error scanning file for relationships: " + item.Name + " - " + e.Message)
		  End Try
		  
		  
		  
		  
		  
		  '//Private Sub ScanFileForRelationships(item As FolderItem)
		  'Try
		  'Var tis As TextInputStream = TextInputStream.Open(item)
		  'Var content As String = tis.ReadAll
		  'tis.Close
		  '
		  '// Clean the code
		  'Var cleanedText As String = CleanCodeForAnalysis(content)
		  '
		  '// Analyze for relationships
		  'AnalyzeFileForRelationships(content)
		  '
		  '// Get all elements and mark them as used
		  'Var allElements() As CodeElement = GetAllElements()
		  'MarkUsedElements(cleanedText, allElements)
		  '
		  'Catch e As IOException
		  'Var errorMsg As String = "Error scanning file for relationships: " + item.Name
		  'Logger.Log(errorMsg)
		  'End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanForRelationships()
		  // Public Sub ScanForRelationships()
		  Logger.Log("=== ScanForRelationships: Analyzing method calls ===")
		  
		  // Get all methods that have been collected
		  Var methods() As CodeElement = GetMethodElements()
		  
		  Logger.Log("Found " + methods.Count.ToString + " methods to analyze for call relationships")
		  
		  // For each method with code, detect what it calls
		  Var methodsWithCode As Integer = 0
		  For Each method As CodeElement In methods
		    If method.Code.Trim <> "" Then
		      methodsWithCode = methodsWithCode + 1
		      DetectMethodCalls(method.Code, method)
		    End If
		  Next
		  
		  Logger.Log("Analyzed " + methodsWithCode.ToString + " methods with code")
		  Logger.Log("=== Relationship scan complete ===")
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProject(folder As FolderItem)
		  Logger.Log("=== ScanProject CALLED ===")
		  
		  // Clear previous data
		  AllElements.RemoveAll
		  ClassElements.RemoveAll
		  MethodElements.RemoveAll
		  ModuleElements.RemoveAll
		  ElementLookup.RemoveAll
		  
		  If folder = Nil Or Not folder.Exists Then Return
		  
		  // First pass: Collect all declarations
		  ScanProjectForDeclarations(folder)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ScanProjectForDeclarations(folder As FolderItem)
		  If folder = Nil Or Not folder.Exists Then Return
		  
		  For Each item As FolderItem In folder.Children
		    Try
		      If item = Nil Then Continue
		      
		      If item.IsFolder Then
		        ScanProjectForDeclarations(item)
		      ElseIf IsXojoSourceFile(item) Then
		        ProcessSourceFile(item)
		      End If
		      
		    Catch e As RuntimeException
		      Logger.Log("Error processing item: " + If(item <> Nil, item.Name, "unknown"))
		      Continue
		    End Try
		  Next
		  
		  // After building relationships and error analysis
		  AnalyzeRefactoringOpportunities()
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SplitParameters(paramList As String) As String()
		  // Private Function SplitParameters(paramList As String) As String()
		  // Split parameter list by commas, handling nested types like Dictionary(String, Integer)
		  
		  Var params() As String
		  Var currentParam As String = ""
		  Var parenDepth As Integer = 0
		  
		  
		  For i As Integer = 0 To paramList.Length 
		    
		    Var c As String = paramList.Middle(i, 1)
		    Select Case c
		    case "("
		      
		      parenDepth = parenDepth + 1
		      'currentParam = currentParam + c
		    case  ")" 
		      parenDepth = parenDepth - 1
		      'currentParam = currentParam + c
		    case "," 
		      // Found a parameter separator at top level
		      If currentParam.Trim <> "" Then
		        params.Add(currentParam.Trim)
		      End If
		      currentParam = ""
		    Else
		      currentParam = currentParam + c
		    End select
		  Next
		  
		  // Add the last parameter
		  If currentParam.Trim <> "" Then
		    params.Add(currentParam.Trim)
		  End If
		  
		  Return params
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		AllElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		ClassElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		DebugMode As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		DetectedSmells() As CodeSmell
	#tag EndProperty

	#tag Property, Flags = &h0
		ElementLookup As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		mBareCatches() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		mElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		MethodElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		mMissingErrorHandling() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		mRiskyOperations() As String
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
End Class
#tag EndClass
