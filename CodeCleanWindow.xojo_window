#tag DesktopWindow
Begin DesktopWindow CodeCleanWindow
   Backdrop        =   0
   BackgroundColor =   &c00F900EB
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   HasTitleBar     =   True
   Height          =   682
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   257478655
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "Scan code and display unused methods and properties"
   Type            =   0
   Visible         =   True
   Width           =   620
   Begin DesktopBevelButton btnScan
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Scan .Xojo_Project File"
      CaptionAlignment=   3
      CaptionDelta    =   0
      CaptionPosition =   1
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      HasBackgroundColor=   False
      Height          =   22
      Icon            =   0
      IconAlignment   =   0
      IconDeltaX      =   0
      IconDeltaY      =   0
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   221
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   37
      Transparent     =   False
      Underline       =   False
      Value           =   False
      Visible         =   True
      Width           =   179
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopTextArea txtResults
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   True
      AllowStyledText =   True
      AllowTabs       =   False
      BackgroundColor =   &cFFFB00D1
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      HasHorizontalScrollbar=   False
      HasVerticalScrollbar=   True
      Height          =   560
      HideSelection   =   True
      Index           =   -2147483648
      Italic          =   False
      Left            =   43
      LineHeight      =   0.0
      LineSpacing     =   1.0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Multiline       =   True
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c000000
      Tooltip         =   ""
      Top             =   95
      Transparent     =   False
      Underline       =   False
      UnicodeMode     =   1
      ValidationMask  =   ""
      Visible         =   True
      Width           =   535
   End
   Begin DesktopImageViewer ImageViewer1
      AllowAutoDeactivate=   True
      AllowTabStop    =   True
      Enabled         =   True
      Height          =   63
      Image           =   1280350207
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   469
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   20
      Transparent     =   False
      Visible         =   True
      Width           =   64
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Sub DebugProjectFiles(folder As FolderItem)
		  'This code Implements a recursive file scanner that processes Xojo project files, 
		  'displaying their content (Or a preview) In a Text area 
		  'while handling potential file access errors gracefully.
		  
		  // Iterate through each item (file or folder) in the current folder's children collection
		  For Each item As FolderItem In folder.Children
		    
		    // Check if the current item is a directory/folder
		    If item.IsFolder Then
		      // Recursively call this same function to process subdirectories
		      // This enables deep scanning of nested folder structures
		      DebugProjectFiles(item)
		      
		      // Check if the current item is a file with specific Xojo project extensions
		      // Using line continuation (_) to split the long conditional across multiple lines
		    ElseIf item.Name.EndsWith(".xojo_code") Or _      // Xojo code files (classes, modules)
		      item.Name.EndsWith(".xojo_window") Or _         // Xojo window/form definitions
		      item.Name.EndsWith(".xojo_form") Or _           // Legacy Xojo form files
		      item.Name.EndsWith(".xojo_project") Then        // Main Xojo project file
		      
		      // Add a visual separator header to the output showing which file is being processed
		      // This helps distinguish between different files in the debug output
		      txtResults.Text = txtResults.Text + "=== FILE: " + item.Name + " ===" + EndOfLine
		      
		      // Wrap file operations in Try-Catch to handle potential I/O errors gracefully
		      Try
		        // Create a TextInputStream to read the file's content
		        // TextInputStream is specifically for reading text-based files
		        Var tis As TextInputStream = TextInputStream.Open(item)
		        
		        // Read the entire file content into a string variable
		        // ReadAll() loads the complete file into memory at once
		        Var content As String = tis.ReadAll
		        
		        // Always close the stream to free system resources and unlock the file
		        // Important for proper file handle management
		        tis.Close
		        
		        // Limit output to prevent UI overwhelm with very large files
		        // Check if file content exceeds 500 characters
		        If content.Length > 500 Then
		          // For large files, show only first 500 characters plus ellipsis
		          // Left(500) extracts the leftmost 500 characters from the string
		          txtResults.Text = txtResults.Text + content.Left(500) + "..." + EndOfLine + EndOfLine
		        Else
		          // For smaller files, display the complete content
		          // Two EndOfLine characters create visual spacing between files
		          txtResults.Text = txtResults.Text + content + EndOfLine + EndOfLine
		        End If
		        
		        // Catch block handles IOException which can occur during file operations
		        // Common causes: file locked, permissions denied, file deleted during read, etc.
		      Catch e As IOException
		        // Display error message instead of crashing the application
		        // Include the actual error message from the exception for debugging
		        txtResults.Text = txtResults.Text + "Error reading file: " + e.Message + EndOfLine + EndOfLine
		      End Try
		      
		      // End of the file type checking conditional
		    End If
		    
		    // End of the For Each loop - continues to next item in folder.Children
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractMethodName(line As String) As String
		  // Private Function ExtractMethodName(line As String) As String
		  // Remove visibility keywords and Sub/Function
		  line = line.Replace("Protected ", "")
		  line = line.Replace("Private ", "")  
		  line = line.Replace("Public ", "")
		  line = line.Replace("Sub ", "")
		  line = line.Replace("Function ", "")
		  
		  // Get name before opening parenthesis or space
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
		Private Function ExtractPropertyName(line As String) As String
		  // Private Function ExtractPropertyName(line As String) As String
		  // Remove visibility keywords and Property
		  line = line.Replace("Protected Property ", "")
		  line = line.Replace("Private Property ", "")
		  line = line.Replace("Public Property ", "")
		  line = line.Replace("Property ", "")
		  
		  // Get name before space or end of line
		  Var spacePos As Integer = line.IndexOf(" ")
		  If spacePos > 0 Then
		    Return line.Left(spacePos).Trim
		  End If
		  
		  Return line.Trim
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractVariableName(line As String) As String
		  // This function handles various Xojo variable declaration formats like:
		  '
		  'Var myVariable As String
		  'Dim count = 0
		  'Var items() As String
		  'myVar As Integer = 42
		  '
		  'The cascading If-statements ensure it finds the earliest termination point To correctly isolate 
		  'just the variable name from the declaration syntax.
		  
		  
		  // Function ExtractVariableName(line As String) As String
		  // Purpose: Parse variable declaration lines to extract clean variable names
		  
		  // STEP 1: Remove variable declaration keywords from the line
		  // This normalizes variable declarations by stripping Xojo's declaration syntax
		  
		  line = line.Replace("Dim ", "")     // Remove "Dim" keyword (traditional declaration style)
		  line = line.Replace("Var ", "")     // Remove "Var" keyword (modern API 2.0 declaration style)
		  
		  // After removal, line contains: "variableName As DataType" or "variableName = value" or variations
		  
		  // STEP 2: Find all possible termination points for the variable name
		  // Variable names can end at several different syntax elements
		  
		  // Find first space character - may indicate type declaration or assignment
		  Var spacePos As Integer = line.IndexOf(" ")
		  
		  // Find " As " keyword specifically - indicates explicit type declaration
		  // This is more precise than just finding any space
		  Var asPos As Integer = line.IndexOf(" As ")
		  
		  // Find equals sign - indicates variable initialization with assignment
		  Var equalPos As Integer = line.IndexOf("=")
		  
		  // Initialize ending position to full line length as fallback
		  // Handles edge cases where none of the terminators are found
		  Var endPos As Integer = line.Length
		  
		  // STEP 3: Determine the earliest valid termination point
		  // Use cascading logic to find where the variable name actually ends
		  
		  // If any space is found, use it as initial boundary
		  If spacePos > 0 Then endPos = spacePos
		  
		  // If " As " is found and comes before current endPos, use it instead
		  // This handles: "myVar As String" -> prefer " As " over generic space
		  If asPos > 0 And asPos < endPos Then endPos = asPos
		  
		  // If equals sign is found and comes before current endPos, use it instead  
		  // This handles: "myVar = 42" -> variable name ends at equals sign
		  If equalPos > 0 And equalPos < endPos Then endPos = equalPos
		  
		  // STEP 4: Extract the variable name using the determined endpoint
		  If endPos > 0 Then
		    // Extract substring from beginning to endPos
		    // Left(endPos) takes characters from position 0 to endPos-1
		    // Trim() removes any residual whitespace from extraction
		    Return line.Left(endPos).Trim
		  End If
		  
		  // STEP 5: Handle malformed input gracefully
		  // Return empty string if no valid variable name could be parsed
		  // This prevents crashes on unexpected input formats
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsSystemMethod(methodName As String) As Boolean
		  // Private Function IsSystemMethod(methodName As String) As Boolean
		  // Skip common system/event methods that are automatically called by Xojo
		  Var systemMethods() As String
		  systemMethods.Add("Constructor")
		  systemMethods.Add("Destructor")
		  systemMethods.Add("Open")
		  systemMethods.Add("Close")
		  systemMethods.Add("Pressed")
		  systemMethods.Add("Action")
		  systemMethods.Add("Paint")
		  systemMethods.Add("Resized")
		  systemMethods.Add("Moved")
		  systemMethods.Add("Activated")
		  systemMethods.Add("Deactivated")
		  systemMethods.Add("KeyDown")
		  systemMethods.Add("MouseDown")
		  systemMethods.Add("MouseUp")
		  systemMethods.Add("MouseMove")
		  systemMethods.Add("Run")
		  
		  For Each sysMethod As String In systemMethods
		    If methodName = sysMethod Then
		      Return True
		    End If
		  Next
		  
		  Return False
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProjectForDeclarations(folder As FolderItem, ByRef results() As String)
		  For Each item As FolderItem In folder.Children
		    If item.IsFolder Then
		      ScanProjectForDeclarations(item, results)
		    ElseIf item.Name.EndsWith(".xojo_code") Or _
		      item.Name.EndsWith(".xojo_window") Or _
		      item.Name.EndsWith(".xojo_form") Then
		      
		      Try
		        Var tis As TextInputStream = TextInputStream.Open(item)
		        Var content As String = tis.ReadAll
		        tis.Close
		        
		        Var lines() As String = content.Split(EndOfLine)
		        Var inMethodOrFunction As Boolean = False
		        
		        For Each line As String In lines
		          line = line.Trim
		          
		          // ====== CONSTANTS ======
		          If line.BeginsWith("#tag Constant, Name = ") Then
		            Var parts() As String = line.Split("Name = ")
		            If parts.LastIndex >= 1 Then
		              Var namepart As String = parts(1)
		              // Extract name before comma
		              Var commaPos As Integer = namepart.IndexOf(",")
		              If commaPos > 0 Then
		                Var constName As String = namepart.Left(commaPos).Trim
		                If results.IndexOf("CONSTANT: " + constName) = -1 Then
		                  results.Add("CONSTANT: " + constName)
		                  System.DebugLog("Found constant: " + constName + " in " + item.Name)
		                End If
		              End If
		            End If
		            
		            // ====== CLASSES ======
		          ElseIf line.Contains("Class ") And (line.BeginsWith("Protected Class ") Or line.BeginsWith("Public Class ") Or line.BeginsWith("Private Class ")) Then
		            Var className As String = ""
		            If line.BeginsWith("Protected Class ") Then
		              className = line.Replace("Protected Class ", "")
		            ElseIf line.BeginsWith("Public Class ") Then
		              className = line.Replace("Public Class ", "")
		            ElseIf line.BeginsWith("Private Class ") Then
		              className = line.Replace("Private Class ", "")
		            End If
		            
		            // Skip system classes that are always "used"
		            If className <> "" And className <> "App" And results.IndexOf("CLASS: " + className) = -1 Then
		              results.Add("CLASS: " + className)
		              System.DebugLog("Found class: " + className + " in " + item.Name)
		            End If
		            
		            // ====== METHODS AND FUNCTIONS ======
		          ElseIf line.BeginsWith("Sub ") Or line.BeginsWith("Function ") Or _
		            line.BeginsWith("Private Sub ") Or line.BeginsWith("Private Function ") Or _
		            line.BeginsWith("Public Sub ") Or line.BeginsWith("Public Function ") Or _
		            line.BeginsWith("Protected Sub ") Or line.BeginsWith("Protected Function ") Then
		            
		            Var methodName As String = ExtractMethodName(line)
		            // Skip system/event methods that are always "used"
		            If methodName <> "" And Not IsSystemMethod(methodName) And results.IndexOf("METHOD: " + methodName) = -1 Then
		              results.Add("METHOD: " + methodName)
		              System.DebugLog("Found method: " + methodName + " in " + item.Name)
		            End If
		            inMethodOrFunction = True
		            
		            // ====== PROPERTIES (look for property-like patterns) ======
		          ElseIf line.BeginsWith("Property ") Or line.BeginsWith("Private Property ") Or  _
		            line.BeginsWith("Public Property ") Or line.BeginsWith("Protected Property ") Then
		            
		            Var propName As String = ExtractPropertyName(line)
		            If propName <> "" And results.IndexOf("PROPERTY: " + propName) = -1 Then
		              results.Add("PROPERTY: " + propName)
		              System.DebugLog("Found property: " + propName + " in " + item.Name)
		            End If
		            
		            // ====== COMPUTED PROPERTIES ======
		          ElseIf line.BeginsWith("#tag ComputedProperty") Then
		            // Look for the actual property declaration in subsequent lines
		            inMethodOrFunction = True
		            
		            // ====== EVENTS ======
		          ElseIf line.BeginsWith("#tag Event") Then
		            // Events are typically followed by Sub EventName
		            inMethodOrFunction = True
		            
		            // ====== VARIABLES (Dim, Var statements at class level) ======
		          ElseIf (line.BeginsWith("Dim ") Or line.BeginsWith("Var ")) And Not inMethodOrFunction Then
		            Var varName As String = ExtractVariableName(line)
		            If varName <> "" And results.IndexOf("VARIABLE: " + varName) = -1 Then
		              results.Add("VARIABLE: " + varName)
		              System.DebugLog("Found variable: " + varName + " in " + item.Name)
		            End If
		            
		            // ====== END TAGS - Reset context ======
		          ElseIf line.BeginsWith("#tag End") Or line = "End" Then
		            inMethodOrFunction = False
		          End If
		        Next
		        
		      Catch e As IOException
		        MessageBox("Error reading file: " + item.NativePath)
		      End Try
		    End If
		  Next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProjectForUsages(folder As FolderItem, declared() As String, ByRef results() As String)
		  For Each item As FolderItem In folder.Children
		    If item.IsFolder Then
		      ScanProjectForUsages(item, declared, results)
		    ElseIf item.Name.EndsWith(".xojo_code") Or _
		      item.Name.EndsWith(".xojo_window") Or _
		      item.Name.EndsWith(".xojo_form") Or _
		      item.Name.EndsWith(".xojo_menu") Then  // Include menu files
		      Try
		        Var tis As TextInputStream = TextInputStream.Open(item)
		        Var originalText As String = tis.ReadAll
		        tis.Close
		        
		        // Strip comments
		        Var rx1 As New RegEx
		        rx1.SearchPattern = "'[^\r\n]*"
		        rx1.ReplacementPattern = ""
		        Var cleanedText As String = rx1.Replace(originalText)
		        
		        // Strip double-quoted strings  
		        Var rx2 As New RegEx
		        rx2.SearchPattern = Chr(34) + ".*?" + Chr(34)
		        rx2.ReplacementPattern = ""
		        cleanedText = rx2.Replace(cleanedText)
		        
		        // Check each declared item for usage
		        For Each fullItem As String In declared
		          // Extract just the name part (remove TYPE: prefix)
		          Var parts() As String = fullItem.Split(": ")
		          Var itemType As String = ""
		          Var word As String = fullItem
		          
		          If parts.LastIndex >= 1 Then
		            itemType = parts(0)
		            word = parts(1)
		          End If
		          
		          Var found As Boolean = False
		          Var whereFound As String = ""
		          
		          // Enhanced usage detection with location tracking
		          Select Case itemType
		          Case "CONSTANT"
		            // Constants can be used in many ways, including menu definitions
		            If cleanedText.IndexOf(word) >= 0 Then
		              found = True
		              whereFound = "anywhere in " + item.Name
		            End If
		            
		          Case "METHOD"
		            // Methods called with parentheses
		            If cleanedText.IndexOf(word + "(") >= 0 Then
		              found = True
		              whereFound = "called in " + item.Name
		              // Also check for method references without parentheses (like AddressOf)
		            ElseIf cleanedText.IndexOf("AddressOf " + word) >= 0 Then
		              found = True
		              whereFound = "referenced via AddressOf in " + item.Name
		              // Check if it appears in other contexts (but be more careful)
		            ElseIf cleanedText.IndexOf(" " + word + " ") >= 0 Then
		              // Additional check: make sure it's not just part of a longer identifier
		              // This might be causing false positives
		              found = True
		              whereFound = "possibly referenced in " + item.Name + " (needs verification)"
		            End If
		            
		          Case "CLASS_VAR"
		            // Class-level variables (properties)
		            If cleanedText.IndexOf("." + word) >= 0 Or _
		              cleanedText.IndexOf(" " + word + " ") >= 0 Or _
		              cleanedText.IndexOf("=" + word) >= 0 Or _
		              cleanedText.IndexOf(word + " =") >= 0 Or _
		              cleanedText.IndexOf(word + ".") >= 0 Then
		              found = True
		            End If
		            
		          Case "LOCAL_VAR"
		            // Simplified local variable detection - avoid the complex method boundary parsing
		            Var varParts() As String = word.Split(".")
		            If varParts.LastIndex >= 1 Then
		              Var methodName As String = varParts(0)
		              Var varName As String = varParts(1)
		              
		              // Look for variable usage anywhere in the file (simpler approach)
		              If cleanedText.IndexOf(" " + varName + " ") >= 0 Or _
		                cleanedText.IndexOf("=" + varName) >= 0 Or _
		                cleanedText.IndexOf(varName + " =") >= 0 Or _
		                cleanedText.IndexOf(varName + ".") >= 0 Or _
		                cleanedText.IndexOf("(" + varName + ")") >= 0 Then
		                
		                // Make sure it's not just the declaration line
		                If Not (cleanedText.IndexOf("Var " + varName + " As") >= 0 Or _
		                  cleanedText.IndexOf("Dim " + varName + " As") >= 0) Then
		                  found = True
		                End If
		              End If
		            End If
		            
		          Case Else
		            // Generic detection for other types
		            If cleanedText.IndexOf(word + "(") >= 0 Or _
		              cleanedText.IndexOf("." + word) >= 0 Or _
		              cleanedText.IndexOf(" " + word + " ") >= 0 Or _
		              cleanedText.IndexOf("=" + word) >= 0 Or _
		              cleanedText.IndexOf(word + "=") >= 0 Then
		              found = True
		              whereFound = "detected in " + item.Name
		            End If
		          End Select
		          
		          If found And results.IndexOf(fullItem) = -1 Then
		            results.Add(fullItem)
		            System.DebugLog("Found usage of: " + fullItem + " - " + whereFound)
		          End If
		        Next
		        
		      Catch e As IOException
		        MessageBox("Error reading: " + item.NativePath)
		      End Try
		    End If
		  Next
		End Sub
	#tag EndMethod


#tag EndWindowCode

#tag Events btnScan
	#tag Event
		Sub Pressed()
		  
		  Var projectFile As FolderItem = FolderItem.ShowOpenFileDialog(".xojo_project")
		  If projectFile = Nil Then
		    txtResults.Text = "No file selected."
		    Return
		  End If
		  
		  txtResults.Text = "Scanning project: " + projectFile.Name + EndOfLine + EndOfLine
		  
		  Var projectFolder As FolderItem = projectFile.Parent
		  
		  Var declaredItems() As String
		  Var usedItems() As String
		  
		  // Scan for declarations and usage
		  ScanProjectForDeclarations(projectFolder, declaredItems)
		  
		  // DEBUG: Show what was actually found
		  txtResults.Text = txtResults.Text + "=== DECLARATIONS FOUND ===" + EndOfLine
		  For Each item As String In declaredItems
		    txtResults.Text = txtResults.Text + item + EndOfLine
		  Next
		  txtResults.Text = txtResults.Text + EndOfLine
		  
		  ScanProjectForUsages(projectFolder, declaredItems, usedItems)
		  
		  // DEBUG: Show what usages were found
		  txtResults.Text = txtResults.Text + "=== USAGES FOUND ===" + EndOfLine
		  For Each item As String In usedItems
		    txtResults.Text = txtResults.Text + item + EndOfLine
		  Next
		  txtResults.Text = txtResults.Text + EndOfLine
		  
		  // Continue with regular analysis...
		  Var unusedMethods() As String
		  Var unusedProperties() As String
		  Var unusedConstants() As String
		  Var unusedClasses() As String
		  Var unusedVariables() As String
		  
		  For Each item As String In declaredItems
		    If usedItems.IndexOf(item) = -1 Then
		      If item.BeginsWith("METHOD:") Then
		        unusedMethods.Add(item)
		      ElseIf item.BeginsWith("PROPERTY:") Then
		        unusedProperties.Add(item)
		      ElseIf item.BeginsWith("CONSTANT:") Then
		        unusedConstants.Add(item)
		      ElseIf item.BeginsWith("CLASS:") Then
		        unusedClasses.Add(item)
		      ElseIf item.BeginsWith("VARIABLE:") Then
		        unusedVariables.Add(item)
		      End If
		    End If
		  Next
		  
		  // Build categorized output
		  Var output As String = "FINAL ANALYSIS RESULTS" + EndOfLine
		  output = output + "Total items scanned: " + declaredItems.Count.ToString + EndOfLine + EndOfLine
		  
		  Var totalUnused As Integer = unusedMethods.Count + unusedProperties.Count + unusedConstants.Count + unusedClasses.Count + unusedVariables.Count
		  
		  If totalUnused = 0 Then
		    output = output + "No unused items found" + EndOfLine
		  Else
		    output = output + "POTENTIALLY UNUSED ITEMS:" + EndOfLine + EndOfLine
		    
		    If unusedMethods.Count > 0 Then
		      output = output + "UNUSED METHODS/FUNCTIONS (" + unusedMethods.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedMethods
		        output = output + "  • " + item.Replace("METHOD: ", "") + EndOfLine
		      Next
		      output = output + EndOfLine
		    End If
		    
		    If unusedConstants.Count > 0 Then
		      output = output + "UNUSED CONSTANTS (" + unusedConstants.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedConstants
		        output = output + "  • " + item.Replace("CONSTANT: ", "") + EndOfLine
		      Next
		      output = output + EndOfLine
		    End If
		    
		    If unusedProperties.Count > 0 Then
		      output = output + "UNUSED PROPERTIES (" + unusedProperties.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedProperties
		        output = output + "  • " + item.Replace("PROPERTY: ", "") + EndOfLine
		      Next
		      output = output + EndOfLine
		    End If
		    
		    If unusedClasses.Count > 0 Then
		      output = output + "UNUSED CLASSES (" + unusedClasses.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedClasses
		        output = output + "  • " + item.Replace("CLASS: ", "") + EndOfLine
		      Next
		      output = output + EndOfLine
		    End If
		    
		    If unusedVariables.Count > 0 Then
		      output = output + "UNUSED VARIABLES (" + unusedVariables.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedVariables
		        output = output + "  • " + item.Replace("VARIABLE: ", "") + EndOfLine
		      Next
		    End If
		  End If
		  
		  txtResults.Text = txtResults.Text + output
		  
		End Sub
	#tag EndEvent
#tag EndEvents
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
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
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
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Autolog"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasTitleBar"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
