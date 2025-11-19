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

	#tag Method, Flags = &h0
		Sub AnalyzeErrorHandling()
		  // Analyze error handling patterns in all methods
		  
		  System.DebugLog("=== Starting Error Handling Analysis ===")
		  
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
		  
		  System.DebugLog("=== Error Handling Analysis Complete ===")
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
		  ScanForRelationships(projectFolder)
		  
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
		End Sub
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
		Private Sub DetectMethodCalls(code As String, callingMethod As CodeElement)
		  // Public Sub DetectMethodCalls(code As String, callingMethod As CodeElement)
		  System.DebugLog("=== DetectMethodCalls DEBUG ===")
		  System.DebugLog("Caller: " + callingMethod.Name)
		  System.DebugLog("Code length: " + code.Length.ToString)
		  
		  If code.Trim = "" Then
		    System.DebugLog("  ⚠️ Code is EMPTY!")
		    Return
		  End If
		  
		  // Show first 200 chars of code
		  Var preview As String = code.Left(Min(200, code.Length))
		  System.DebugLog("  Preview: " + preview)
		  
		  // Get all methods to check against
		  Var methods() As CodeElement = GetMethodElements()
		  
		  // Check all known methods to see if they're called in this code
		  For Each methodElement As CodeElement In methods
		    Var methodName As String = methodElement.Name
		    
		    // Look for method calls: MethodName( or object.MethodName(
		    If code.IndexOf(methodName + "(") >= 0 Then
		      // Found a call - create relationship if not already exists
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
		        System.DebugLog("  Relationship: " + callingMethod.Name + " calls " + methodElement.Name)
		      End If
		    End If
		  Next
		  
		  System.DebugLog("  Detected " + callingMethod.CallsTo.Count.ToString + " calls")
		  
		  If callingMethod.CallsTo.Count > 0 Then
		    For Each called As CodeElement In callingMethod.CallsTo
		      System.DebugLog("    → " + called.Name)
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
		  Var paramList As String = signature.Mid(openParen + 1, closeParen - openParen+1).Trim
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
		  System.DebugLog("Number of parameters = " + element.ParameterCount.ToString)
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
		  // Private Sub FinalizeMethod(context As ParsingContext)
		  // Save the accumulated code to the method element
		  
		  If Not context.InMethodOrFunction Then
		    Return
		  End If
		  
		  If context.CurrentMethodFullPath = "" Then
		    Return
		  End If
		  
		  // Find the method element and store its code
		  For Each element As CodeElement In mElements
		    If element.FullPath = context.CurrentMethodFullPath And element.ElementType = "METHOD" Then
		      element.Code = context.CurrentMethodCode.Trim
		      Exit
		    End If
		  Next
		  
		  // Reset tracking
		  context.InMethodOrFunction = False
		  context.CurrentMethodFullPath = ""
		  context.CurrentMethodCode = ""
		  
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
		Function GetAllElements() As CodeElement()
		  // Public Function GetAllElements() As CodeElement()
		  Return mElements
		  
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
		Function GetMethodElements() As CodeElement()
		  // Public Function GetMethodElements() As CodeElement()
		  Var methods() As CodeElement
		  
		  For Each element As CodeElement In mElements
		    System.DebugLog("Checking: " + element.Name + " Type: " + element.ElementType)
		    
		    If element.ElementType = "METHOD" Then  // ← Is this matching?
		      methods.Add(element)
		    End If
		  Next
		  
		  System.DebugLog("Found " + methods.Count.ToString + " methods")
		  Return methods
		  
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
		Function GetUnusedElements() As CodeElement()
		  Var unused() As CodeElement
		  
		  For Each element As CodeElement In AllElements
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
		        System.DebugLog("Stored code for method: " + currentMethodFullPath + " (" + currentMethodCode.Length.ToString + " chars)")
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
		Private Sub MarkUsedElements(cleanedText As String)
		  For Each element As CodeElement In AllElements
		    Var name As String = element.Name
		    
		    Select Case element.ElementType
		    Case "METHOD"
		      If cleanedText.IndexOf(name + "(") >= 0 Or _
		        cleanedText.IndexOf("AddressOf " + name) >= 0 Then
		        element.IsUsed = True
		      End If
		      
		    Case "CLASS"
		      If cleanedText.IndexOf("New " + name) >= 0 Or _
		        cleanedText.IndexOf("As " + name) >= 0 Or _
		        cleanedText.IndexOf("Inherits " + name) >= 0 Then
		        element.IsUsed = True
		      End If
		      
		    Case "MODULE"
		      If cleanedText.IndexOf(name + ".") >= 0 Then
		        element.IsUsed = True
		      End If
		      
		    Case "PROPERTY", "VARIABLE"
		      If cleanedText.IndexOf("." + name) >= 0 Or _
		        cleanedText.IndexOf(" " + name + " ") >= 0 Or _
		        cleanedText.IndexOf(name + " =") >= 0 Then
		        element.IsUsed = True
		      End If
		    End Select
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
		  //Private Sub ProcessMethodDeclaration(declaration As String, context As ParsingContext)
		  // Handle method/function declarations
		  
		  // Finalize any previous method that was still open
		  If context.InMethodOrFunction Then
		    FinalizeMethod(context)
		  End If
		  
		  Var methodName As String = ExtractMethodName(declaration)
		  Var fullPath As String = BuildFullPath(context.CurrentModule, context.CurrentClass, methodName)
		  
		  Var element As New CodeElement("METHOD", methodName, fullPath, context.FileName)
		  
		  // Extract parameter information from the signature
		  ExtractParameterInfo(declaration, element)
		  
		  mElements.Add(element)
		  
		  // Start tracking this method
		  context.InMethodOrFunction = True
		  context.CurrentMethodFullPath = fullPath
		  context.CurrentMethodCode = ""
		  
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
		    System.DebugLog("Error reading file: " + item.Name + " - " + e.Message)
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
		      System.DebugLog("    Total elements: " + mElements.Count.ToString)
		      ElementLookup.Value(fullPath) = element
		      System.DebugLog("Found variable: " + fullPath + " in " + fileName)
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
		  // REFACTORED: Now broken into focused helper methods
		  
		  Try
		    Var tis As TextInputStream = TextInputStream.Open(item)
		    Var content As String = tis.ReadAll
		    tis.Close
		    
		    // Clean the code
		    Var cleanedText As String = CleanCodeForAnalysis(content)
		    
		    // Analyze for relationships
		    AnalyzeFileForRelationships(content)
		    
		    // Mark elements as used
		    MarkUsedElements(cleanedText)
		    
		  Catch e As IOException
		    System.DebugLog("Error scanning file for relationships: " + item.Name)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanForRelationships(folder As FolderItem)
		  // Public Sub ScanForRelationships(projectFolder As FolderItem)
		  System.DebugLog("=== ScanForRelationships ===")
		  
		  // Get all methods
		  Var methods() As CodeElement = GetMethodElements()
		  
		  System.DebugLog("Analyzing " + methods.Count.ToString + " methods for relationships...")
		  
		  // For each method with code, detect what it calls
		  For Each method As CodeElement In methods
		    If method.Code.Trim <> "" Then
		      DetectMethodCalls(method.Code, method)
		    End If
		  Next
		  
		  System.DebugLog("=== Relationship scan complete ===")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProject(folder As FolderItem)
		  System.DebugLog("=== ScanProject CALLED ===")
		  
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
		      System.DebugLog("Error processing item: " + If(item <> Nil, item.Name, "unknown"))
		      Continue
		    End Try
		  Next
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
		ElementLookup As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h0
		mElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		MethodElements() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleElements() As CodeElement
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
