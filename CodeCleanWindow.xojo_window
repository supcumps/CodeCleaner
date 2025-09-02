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
      Active          =   False
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
      PanelIndex      =   0
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      Tooltip         =   ""
      Top             =   20
      Transparent     =   False
      Visible         =   True
      Width           =   64
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
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
		  'This Function  Handles the complexity Of Xojo's method declaration syntax by systematically stripping away language keywords 
		  'and access modifiers, then using intelligent boundary detection to isolate just the method name.
		  'The priority logic (parenthesis over space) Is particularly important because it correctly Handles both parameterized methods
		  'like MyMethod(param As String) And parameterless functions With Return types like Calculate As Integer. 
		  'This ensures your unused code scanner can accurately identify method declarations regardless of their specific syntax variation.
		  
		  // Private Function ExtractMethodName(line As String) As String
		  // Purpose: Parse Xojo method/function declaration lines to extract clean method names
		  // Used by CodeCleaner to identify declared methods for unused code analysis
		  
		  // STEP 1: Strip all visibility modifiers and declaration keywords
		  // This normalizes method declarations by removing Xojo's access control syntax
		  // Examples: "Protected Sub MyMethod(" -> "MyMethod("
		  //          "Private Function Calculate(" -> "Calculate("
		  
		  line = line.Replace("Protected ", "")  // Remove protected access modifier
		  line = line.Replace("Private ", "")    // Remove private access modifier  
		  line = line.Replace("Public ", "")     // Remove public access modifier
		  line = line.Replace("Sub ", "")        // Remove subroutine declaration keyword
		  line = line.Replace("Function ", "")   // Remove function declaration keyword
		  
		  // After cleanup, line contains: "MethodName(parameters)" or "MethodName As ReturnType"
		  
		  // STEP 2: Find potential termination points for the method name
		  // Method names can end at different syntax elements depending on declaration style
		  
		  // Find opening parenthesis - indicates start of parameter list
		  // Example: "MyMethod(param1 As String)" -> parenPos = 8
		  Var parenPos As Integer = line.IndexOf("(")
		  
		  // Find first space character - may indicate return type declaration or other syntax
		  // Example: "Calculate As Integer" -> spacePos = 9
		  Var spacePos As Integer = line.IndexOf(" ")
		  
		  // Initialize ending position to full line length as safety fallback
		  // This handles edge cases where neither parenthesis nor space are found
		  Var endPos As Integer = line.Length
		  
		  // STEP 3: Determine the correct termination point using priority logic
		  // Parenthesis takes precedence over space when both are present
		  
		  // If parenthesis found AND (no space found OR parenthesis comes before space)
		  // This handles: "MyMethod(params)" and "MyMethod(params) As ReturnType"
		  If parenPos > 0 And (spacePos = -1 Or parenPos < spacePos) Then
		    endPos = parenPos  // Method name ends at opening parenthesis
		    
		    // Otherwise, if space is found, use it as the boundary
		    // This handles: "MyMethod As ReturnType" (function without explicit parameters)
		  ElseIf spacePos > 0 Then
		    endPos = spacePos  // Method name ends at first space
		  End If
		  
		  // STEP 4: Extract and return the clean method name
		  If endPos > 0 Then
		    // Extract substring from beginning to endPos
		    // Left(endPos) takes characters from position 0 to endPos-1
		    // Trim() removes any residual whitespace from the extraction process
		    Return line.Left(endPos).Trim
		  End If
		  
		  // STEP 5: Handle malformed input gracefully
		  // Return empty string if no valid method name could be parsed
		  // This prevents the CodeCleaner from crashing on unexpected declaration formats
		  Return ""
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ExtractPropertyName(line As String) As String
		  ' Property names always End at the first space (which typically precedes the "As DataType" clause) 
		  'Or represent the entire remaining String If no space Is found.
		  'The Function Handles both explicitly typed properties (Property MyValue As String) And simple Property declarations. 
		  'This ensures  CodeCleaner can accurately identify Property declarations regardless Of whether they include type information, 
		  'making the unused code detection more comprehensive.
		  
		  
		  // Private Function ExtractPropertyName(line As String) As String
		  // Purpose: Parse Xojo property declaration lines to extract clean property names
		  // Used by CodeCleaner to identify declared properties for unused code analysis
		  
		  // STEP 1: Strip all visibility modifiers and property declaration keywords
		  // This normalizes property declarations by removing Xojo's access control syntax
		  // Examples: "Protected Property MyValue As String" -> "MyValue As String"
		  //          "Private Property Count" -> "Count"
		  //          "Property IsEnabled As Boolean" -> "IsEnabled As Boolean"
		  
		  line = line.Replace("Protected Property ", "")  // Remove protected access modifier + Property keyword
		  line = line.Replace("Private Property ", "")    // Remove private access modifier + Property keyword
		  line = line.Replace("Public Property ", "")     // Remove public access modifier + Property keyword
		  line = line.Replace("Property ", "")            // Remove standalone Property keyword (default visibility)
		  
		  // After cleanup, line contains: "PropertyName As DataType" or just "PropertyName"
		  
		  // STEP 2: Find the boundary where the property name ends
		  // Property names typically end at the first space, which indicates type declaration
		  // Example: "MyValue As String" -> property name is "MyValue"
		  
		  Var spacePos As Integer = line.IndexOf(" ")
		  
		  // STEP 3: Extract property name based on space detection
		  If spacePos > 0 Then
		    // Space found - property name ends before the space
		    // This handles: "PropertyName As DataType" declarations
		    // Left(spacePos) extracts characters from position 0 to spacePos-1
		    // Trim() removes any residual whitespace
		    Return line.Left(spacePos).Trim
		  End If
		  
		  // STEP 4: Handle property declarations without explicit type
		  // No space found - entire remaining string is the property name
		  // This handles simple property declarations like just "PropertyName"
		  // or cases where the type declaration was already stripped
		  // Trim() removes leading/trailing whitespace from the full string
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
		  'This Function Is crucial For the accuracy Of the CodeCleaner tool. 
		  'It acts As a "whitelist" filter that prevents False positives by recognizing methods that are automatically called by the Xojo framework, 
		  'even though no explicit calls To them appear In your source code.
		  'Without this Function, the code scanner would incorrectly flag essential Event handlers like Paint, 
		  'Pressed, Or Constructor As unused, when In reality they're invoked automatically by Xojo's runtime system. 
		  'This makes the code analysis much more reliable by focusing only on truly user-defined methods that might genuinely be unused.
		  
		  // Private Function IsSystemMethod(methodName As String) As Boolean
		  // Purpose: Identify Xojo framework methods that should be excluded from unused code analysis
		  // These are event handlers and lifecycle methods automatically called by the Xojo runtime
		  // Prevents false positives in CodeCleaner by recognizing "implicitly used" methods
		  
		  // STEP 1: Build comprehensive list of Xojo system/framework method names
		  // These methods are automatically invoked by the Xojo runtime and should never
		  // be flagged as "unused" even if no explicit calls are found in the code
		  
		  Var systemMethods() As String
		  
		  // Object lifecycle methods - automatically called during object creation/destruction
		  systemMethods.Add("Constructor")    // Called when object is instantiated
		  systemMethods.Add("Destructor")     // Called when object is being destroyed/garbage collected
		  
		  // Window/Control lifecycle events - automatically triggered by the UI framework
		  systemMethods.Add("Open")           // Called when window/control first opens/appears
		  systemMethods.Add("Close")          // Called when window/control is being closed/hidden
		  
		  // User interface event handlers - automatically invoked by user interactions
		  systemMethods.Add("Pressed")        // Button press events
		  systemMethods.Add("Action")         // Generic action events (menus, buttons, etc.)
		  
		  // Window drawing and layout events - automatically called during UI rendering
		  systemMethods.Add("Paint")          // Called when control needs to redraw itself
		  systemMethods.Add("Resized")        // Called when window/control changes size
		  systemMethods.Add("Moved")          // Called when window/control changes position
		  
		  // Window state change events - automatically triggered by window focus changes
		  systemMethods.Add("Activated")      // Called when window gains focus/becomes active
		  systemMethods.Add("Deactivated")    // Called when window loses focus/becomes inactive
		  
		  // Input event handlers - automatically invoked by keyboard/mouse interactions
		  systemMethods.Add("KeyDown")        // Called when user presses a key
		  systemMethods.Add("MouseDown")      // Called when user presses mouse button
		  systemMethods.Add("MouseUp")        // Called when user releases mouse button
		  systemMethods.Add("MouseMove")      // Called when user moves mouse over control
		  
		  // Application lifecycle methods - automatically called by Xojo runtime
		  systemMethods.Add("Run")            // Called when application starts (App.Run event)
		  
		  // STEP 2: Check if the provided method name matches any system method
		  // Use exact string comparison to determine if this is a framework method
		  
		  For Each sysMethod As String In systemMethods
		    If methodName = sysMethod Then
		      // Method name found in system methods list
		      // Return True to indicate this should be excluded from unused analysis
		      Return True
		    End If
		  Next
		  
		  // STEP 3: Method not found in system methods list
		  // Return False to indicate this is a user-defined method that should be
		  // included in unused code analysis
		  Return False
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProjectForDeclarations(folder As FolderItem, ByRef results() As String)
		  '
		  'This Is the core parsing logic Of your CodeCleaner application. 
		  'This method performs sophisticated Static analysis Of Xojo project files by:
		  '
		  'Recursively traversing the project directory Structure To find all relevant files
		  'Using context awareness (the inMethodOrFunction flag) To distinguish between Class-level declarations And local variables
		  'Implementing pattern matching For different Xojo syntax constructs (constants, classes, methods, properties, variables)
		  'Preventing duplicates by checking If items already exist In the results Array
		  'Filtering out System artifacts like the "App" Class And System methods that would create False positives
		  '
		  'The context tracking ensures that local variables declared inside methods aren't mistakenly flagged as unused class members. 
		  'This makes your analysis much more accurate by focusing only on Class-level declarations that could genuinely be unused.
		  '
		  For Each item As FolderItem In folder.Children
		    // RECURSIVE DIRECTORY TRAVERSAL
		    // Process each file and subdirectory in the current folder
		    If item.IsFolder Then
		      // Recursively scan subdirectories to handle nested project structures
		      // This ensures all Xojo project files are found regardless of folder depth
		      ScanProjectForDeclarations(item, results)
		      
		      // FILE TYPE FILTERING
		      // Only process Xojo project files by checking file extensions
		    ElseIf item.Name.EndsWith(".xojo_code") Or _      // Class modules, regular modules
		      item.Name.EndsWith(".xojo_window") Or _         // Window definitions (forms)
		      item.Name.EndsWith(".xojo_form") Then           // Legacy form files
		      
		      // FILE READING WITH ERROR HANDLING
		      Try
		        // Open file as text stream for reading Xojo project file content
		        Var tis As TextInputStream = TextInputStream.Open(item)
		        Var content As String = tis.ReadAll  // Load entire file content into memory
		        tis.Close  // Always close stream to release file handle
		        
		        // CONTENT PARSING PREPARATION
		        // Split file content into individual lines for line-by-line analysis
		        Var lines() As String = content.Split(EndOfLine)
		        
		        // Context tracking variable - determines if we're inside a method/function body
		        // This prevents treating local variables as class-level declarations
		        Var inMethodOrFunction As Boolean = False
		        
		        // LINE-BY-LINE PARSING LOOP
		        For Each line As String In lines
		          line = line.Trim  // Remove leading/trailing whitespace for consistent parsing
		          
		          // ====== CONSTANTS DETECTION ======
		          // Parse Xojo constant declarations: "#tag Constant, Name = ConstantName, Value = ..."
		          If line.BeginsWith("#tag Constant, Name = ") Then
		            Var parts() As String = line.Split("Name = ")
		            If parts.LastIndex >= 1 Then
		              Var namepart As String = parts(1)  // Get everything after "Name = "
		              
		              // Extract constant name (stops at comma before other attributes)
		              // Example: "ConstantName, Type = Integer" -> extract "ConstantName"
		              Var commaPos As Integer = namepart.IndexOf(",")
		              If commaPos > 0 Then
		                Var constName As String = namepart.Left(commaPos).Trim
		                
		                // Avoid duplicate entries in results array
		                If results.IndexOf("CONSTANT: " + constName) = -1 Then
		                  results.Add("CONSTANT: " + constName)
		                  System.DebugLog("Found constant: " + constName + " in " + item.Name)
		                End If
		              End If
		            End If
		            
		            // ====== CLASSES DETECTION ======
		            // Parse class declarations with visibility modifiers
		            // Must contain "Class " and begin with a visibility keyword to avoid false matches
		          ElseIf line.Contains("Class ") And (line.BeginsWith("Protected Class ") Or line.BeginsWith("Public Class ") Or line.BeginsWith("Private Class ")) Then
		            Var className As String = ""
		            
		            // Strip visibility keywords to extract clean class name
		            If line.BeginsWith("Protected Class ") Then
		              className = line.Replace("Protected Class ", "")
		            ElseIf line.BeginsWith("Public Class ") Then
		              className = line.Replace("Public Class ", "")
		            ElseIf line.BeginsWith("Private Class ") Then
		              className = line.Replace("Private Class ", "")
		            End If
		            
		            // Filter out system classes and avoid duplicates
		            // "App" is excluded because it's always implicitly used by Xojo runtime
		            If className <> "" And className <> "App" And results.IndexOf("CLASS: " + className) = -1 Then
		              results.Add("CLASS: " + className)
		              System.DebugLog("Found class: " + className + " in " + item.Name)
		            End If
		            
		            // ====== METHODS AND FUNCTIONS DETECTION ======
		            // Comprehensive parsing of method/function declarations with all visibility combinations
		          ElseIf line.BeginsWith("Sub ") Or line.BeginsWith("Function ") Or _
		            line.BeginsWith("Private Sub ") Or line.BeginsWith("Private Function ") Or _
		            line.BeginsWith("Public Sub ") Or line.BeginsWith("Public Function ") Or _
		            line.BeginsWith("Protected Sub ") Or line.BeginsWith("Protected Function ") Then
		            
		            // Use helper function to extract clean method name
		            Var methodName As String = ExtractMethodName(line)
		            
		            // Filter out system event methods and avoid duplicates
		            // System methods (Constructor, Paint, etc.) are excluded because they're automatically called
		            If methodName <> "" And Not IsSystemMethod(methodName) And results.IndexOf("METHOD: " + methodName) = -1 Then
		              results.Add("METHOD: " + methodName)
		              System.DebugLog("Found method: " + methodName + " in " + item.Name)
		            End If
		            
		            // Update context - we're now inside a method/function body
		            inMethodOrFunction = True
		            
		            // ====== PROPERTIES DETECTION ======
		            // Parse property declarations with all visibility modifiers
		          ElseIf line.BeginsWith("Property ") Or line.BeginsWith("Private Property ") Or  _
		            line.BeginsWith("Public Property ") Or line.BeginsWith("Protected Property ") Then
		            
		            // Use helper function to extract clean property name
		            Var propName As String = ExtractPropertyName(line)
		            If propName <> "" And results.IndexOf("PROPERTY: " + propName) = -1 Then
		              results.Add("PROPERTY: " + propName)
		              System.DebugLog("Found property: " + propName + " in " + item.Name)
		            End If
		            
		            // ====== COMPUTED PROPERTIES DETECTION ======
		            // Computed properties are defined using special Xojo tags
		          ElseIf line.BeginsWith("#tag ComputedProperty") Then
		            // Mark context change - computed property definitions contain method-like code
		            // The actual property name will be found in subsequent lines
		            inMethodOrFunction = True
		            
		            // ====== EVENTS DETECTION ======
		            // Event handlers are defined using Xojo's tag system
		          ElseIf line.BeginsWith("#tag Event") Then
		            // Mark context change - events contain method-like code
		            // Event method names will be detected by the method parsing logic above
		            inMethodOrFunction = True
		            
		            // ====== CLASS-LEVEL VARIABLES DETECTION ======
		            // Only detect variables at class level (not inside methods/functions)
		            // This prevents local variables from being flagged as unused class members
		          ElseIf (line.BeginsWith("Dim ") Or line.BeginsWith("Var ")) And Not inMethodOrFunction Then
		            // Use helper function to extract clean variable name
		            Var varName As String = ExtractVariableName(line)
		            If varName <> "" And results.IndexOf("VARIABLE: " + varName) = -1 Then
		              results.Add("VARIABLE: " + varName)
		              System.DebugLog("Found variable: " + varName + " in " + item.Name)
		            End If
		            
		            // ====== CONTEXT RESET ======
		            // End tags indicate we're leaving method/function/event/property contexts
		          ElseIf line.BeginsWith("#tag End") Or line = "End" Then
		            // Reset context - we're back at class level
		            inMethodOrFunction = False
		          End If
		        Next
		        
		        // ERROR HANDLING
		        // Handle file I/O errors gracefully without crashing the application
		      Catch e As IOException
		        MessageBox("Error reading file: " + item.NativePath)
		      End Try
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProjectForUsages(folder As FolderItem, declared() As String, ByRef results() As String)
		  'This Is the second half Of  CodeCleaner's analysis engine - the usage detection system. 
		  'This method performs the crucial task of determining which declared items are actually referenced in the codebase.
		  'The preprocessing steps (removing comments And String literals) are particularly important because they prevent False positives 
		  'where variable names might appear In comments Or String constants but aren't actually being used as code references.
		  'The type-specific detection logic Is sophisticated:
		  '
		  'Constants use simple substring matching since they can appear anywhere
		  'Methods check For multiple usage patterns (Function calls, AddressOf references, assignments)
		  'Local variables include method context To avoid cross-method confusion
		  'Class variables look For various access patterns (dot notation, assignments, etc.)
		  '
		  'The combination Of this method With ScanProjectForDeclarations creates a comprehensive Static analysis System 
		  'that can accurately identify truly unused code elements In Xojo projects, helping developers clean up their codebase efficiently.
		  '
		  
		  
		  For Each item As FolderItem In folder.Children
		    // RECURSIVE DIRECTORY TRAVERSAL
		    // Process each file and subdirectory in the current folder
		    If item.IsFolder Then
		      // Recursively scan subdirectories to ensure complete project coverage
		      ScanProjectForUsages(item, declared, results)
		      
		      // FILE TYPE FILTERING - Extended to include menu files
		      // Process all Xojo project files that might contain code references
		    ElseIf item.Name.EndsWith(".xojo_code") Or _       // Class modules, regular modules
		      item.Name.EndsWith(".xojo_window") Or _          // Window definitions (forms)
		      item.Name.EndsWith(".xojo_form") Or _            // Legacy form files
		      item.Name.EndsWith(".xojo_menu") Then            // Menu definition files (may reference constants)
		      
		      // FILE READING WITH ERROR HANDLING
		      Try
		        // Read entire file content for comprehensive text analysis
		        Var tis As TextInputStream = TextInputStream.Open(item)
		        Var originalText As String = tis.ReadAll
		        tis.Close
		        
		        // TEXT PREPROCESSING - Clean content to avoid false matches
		        // Check each declared item for usage in string literals BEFORE cleaning
		        For Each fullItem As String In declared
		          // Extract just the name part (remove TYPE: prefix)
		          Var parts() As String = fullItem.Split(": ")
		          Var itemType As String = ""
		          Var word As String = fullItem
		          
		          If parts.LastIndex >= 1 Then
		            itemType = parts(0)
		            word = parts(1)
		          End If
		          
		          // Check for class names referenced as string literals (before text cleaning removes them)
		          // This catches dynamic instantiation, plugin registration, configuration files, etc.
		          If itemType = "CLASS" And originalText.IndexOf(Chr(34) + word + Chr(34)) >= 0 Then
		            If results.IndexOf(fullItem) = -1 Then
		              results.Add(fullItem)
		              System.DebugLog("Found string literal usage of: " + fullItem + " in " + item.Name)
		            End If
		          End If
		        Next
		        // STEP 1: Remove single-line comments to prevent matches within comments
		        // Pattern: '[^\r\n]*' matches apostrophe followed by any characters until line end
		        Var rx1 As New RegEx
		        rx1.SearchPattern = "'[^\r\n]*"      // Match: ' comment text
		        rx1.ReplacementPattern = ""           // Replace with empty string
		        Var cleanedText As String = rx1.Replace(originalText)
		        
		        // STEP 2: Remove string literals to prevent matches within quoted text
		        // Pattern: Chr(34) + ".*?" + Chr(34) matches content between double quotes
		        // Chr(34) is the double quote character (") - using Chr() avoids quote escaping issues
		        Var rx2 As New RegEx
		        rx2.SearchPattern = Chr(34) + ".*?" + Chr(34)  // Match: "any string content"
		        rx2.ReplacementPattern = ""                     // Replace with empty string
		        cleanedText = rx2.Replace(cleanedText)
		        
		        // USAGE DETECTION LOOP
		        // Check each declared item to see if it's used anywhere in this file
		        For Each fullItem As String In declared
		          
		          // PARSE DECLARED ITEM FORMAT
		          // Items are stored as "TYPE: ItemName" - extract both parts
		          Var parts() As String = fullItem.Split(": ")
		          Var itemType As String = ""
		          Var word As String = fullItem  // Fallback to full item if parsing fails
		          
		          If parts.LastIndex >= 1 Then
		            itemType = parts(0)  // Extract type (CONSTANT, METHOD, etc.)
		            word = parts(1)      // Extract the actual name to search for
		          End If
		          
		          // USAGE TRACKING VARIABLES
		          Var found As Boolean = False
		          Var whereFound As String = ""  // Optional location description for debugging
		          
		          // TYPE-SPECIFIC USAGE DETECTION
		          // Different code elements have different usage patterns
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
		              found = True
		              whereFound = "possibly referenced in " + item.Name + " (needs verification)"
		            End If
		            
		          Case "CLASS"
		            // Class instantiation: New ClassName
		            If cleanedText.IndexOf("New " + word) >= 0 Then
		              found = True
		              whereFound = "instantiated in " + item.Name
		              // Type declarations: As ClassName
		            ElseIf cleanedText.IndexOf("As " + word) >= 0 Then
		              found = True
		              whereFound = "used as type in " + item.Name
		              // Inheritance: Inherits ClassName  
		            ElseIf cleanedText.IndexOf("Inherits " + word) >= 0 Then
		              found = True
		              whereFound = "inherited in " + item.Name
		              // Generic patterns (keep existing for other usages)
		            ElseIf cleanedText.IndexOf(" " + word + " ") >= 0 Then
		              found = True
		              whereFound = "referenced in " + item.Name
		            ElseIf cleanedText.IndexOf("Catch ") >= 0 And cleanedText.IndexOf("As " + word) >= 0 Then
		              found = True
		            ElseIf cleanedText.IndexOf("Raise New " + word) >= 0 Then
		              found = True
		            End If
		            
		          Case "PROPERTY", "VARIABLE"
		            // Class-level variables and properties
		            If cleanedText.IndexOf("." + word) >= 0 Or _
		              cleanedText.IndexOf(" " + word + " ") >= 0 Or _
		              cleanedText.IndexOf("=" + word) >= 0 Or _
		              cleanedText.IndexOf(word + " =") >= 0 Or _
		              cleanedText.IndexOf(word + ".") >= 0 Then
		              found = True
		              whereFound = "property/variable usage in " + item.Name
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
		                  whereFound = "local variable usage in " + item.Name
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
		          
		          // RECORD USAGE RESULTS
		          // Add to results array if usage found and not already recorded
		          If found And results.IndexOf(fullItem) = -1 Then
		            results.Add(fullItem)  // Mark this item as "used"
		            System.DebugLog("Found usage of: " + fullItem + " - " + whereFound)
		          End If
		        Next
		        
		        // ERROR HANDLING
		        // Handle file I/O errors gracefully without crashing the scan
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
