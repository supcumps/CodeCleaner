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
      Top             =   20
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
   Begin DesktopBevelButton ExportButton
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Export Results"
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
      Scope           =   2
      TabIndex        =   3
      TabPanelIndex   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   54
      Transparent     =   False
      Underline       =   False
      Value           =   False
      Visible         =   True
      Width           =   179
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Method, Flags = &h21
		Private Function BuildFullName(itemType As String, itemName As String, currentModule As String, currentClass As String) As String
		  // Private Function BuildFullName(itemType As String, itemName As String, currentModule As String, currentClass As String) As String
		  Var fullName As String = itemType + ": "
		  
		  If currentModule <> "" And currentClass <> "" Then
		    fullName = fullName + currentModule + "." + currentClass + "." + itemName
		  ElseIf currentModule <> "" Then
		    fullName = fullName + currentModule + "." + itemName
		  ElseIf currentClass <> "" Then
		    fullName = fullName + currentClass + "." + itemName
		  Else
		    fullName = fullName + itemName
		  End If
		  
		  Return fullName
		  
		End Function
	#tag EndMethod

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
		    ElseIf IsXojoSourceFile(item) Then
		      
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
		Private Function IsClassDeclaration(line As String) As Boolean
		  // Private Function IsClassDeclaration(line As String) As Boolean
		  Return line.Contains("Class ") And _
		  (line.BeginsWith("Protected Class ") Or _
		  line.BeginsWith("Public Class ") Or _
		  line.BeginsWith("Private Class ") Or _
		  line.BeginsWith("Class "))
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsInterfaceDeclaration(line As String) As Boolean
		  // Private Function IsInterfaceDeclaration(line As String) As Boolean
		  Return line.Contains("Interface ") And _
		  (line.BeginsWith("Interface ") Or _
		  line.BeginsWith("Protected Interface ") Or _
		  line.BeginsWith("Public Interface ") Or _
		  line.BeginsWith("Private Interface "))
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsMethodOrFunctionDeclaration(line As String) As Boolean
		  // Private Function IsMethodOrFunctionDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Sub ") Or line.BeginsWith("Function ") Or _
		  line.BeginsWith("Private Sub ") Or line.BeginsWith("Private Function ") Or _
		  line.BeginsWith("Public Sub ") Or line.BeginsWith("Public Function ") Or _
		  line.BeginsWith("Protected Sub ") Or line.BeginsWith("Protected Function ")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsModuleDeclaration(line As String) As Boolean
		  // Private Function IsModuleDeclaration(line As String) As Boolean
		  Return line.Contains("Module ") And _
		  (line.BeginsWith("Module ") Or _
		  line.BeginsWith("Protected Module ") Or _
		  line.BeginsWith("Public Module ") Or _
		  line.BeginsWith("Private Module ") Or _
		  line.BeginsWith("Global Module "))
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsPropertyDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Property ") Or _
		  line.BeginsWith("Private Property ") Or _
		  line.BeginsWith("Public Property ") Or _
		  line.BeginsWith("Protected Property ")
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

	#tag Method, Flags = &h21
		Private Function IsVariableDeclaration(line As String) As Boolean
		  // Private Function IsVariableDeclaration(line As String) As Boolean
		  Return line.BeginsWith("Dim ") Or line.BeginsWith("Var ")
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsXojoSourceFile(item As FolderItem) As Boolean
		  If item = Nil Or Not item.Exists Then Return False
		  
		  Var name As String = item.Name.Lowercase
		  
		  Return name.EndsWith(".xojo_code") Or _
		  name.EndsWith(".xojo_window") Or _
		  name.EndsWith(".xojo_form") Or _
		  name.EndsWith(".xojo_mobile_screen") Or _
		  name.EndsWith(".xojo_mobile_container") Or _
		  name.EndsWith(".xojo_ios_view") Or _
		  name.EndsWith(".xojo_ios_screen") Or _
		  name.EndsWith(".xojo_menu") Or _
		  name.EndsWith(".xojo_toolbar") Or _
		  name.Contains(".xojo_")  // Catch-all for any xojo_ files
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ParseFileContent(content As String, fileName As String, ByRef results() As String)
		  Var lines() As String = content.Split(EndOfLine)
		  
		  // Context tracking variables
		  Var currentModule As String = ""
		  Var currentClass As String = ""
		  Var currentInterface As String = ""
		  Var inMethodOrFunction As Boolean = False
		  
		  For Each line As String In lines
		    line = line.Trim
		    
		    // Skip empty lines
		    If line = "" Then Continue
		    
		    // Process based on line content
		    If line.BeginsWith("#tag Constant") Then
		      ProcessConstantDeclaration(line, currentModule, currentClass, fileName, results)
		      
		    ElseIf IsModuleDeclaration(line) Then
		      currentModule = ProcessModuleDeclaration(line, fileName, results)
		      currentClass = ""  // Reset class when entering module
		      currentInterface = ""
		      
		    ElseIf IsInterfaceDeclaration(line) Then
		      currentInterface = ProcessInterfaceDeclaration(line, currentModule, fileName, results)
		      currentClass = currentInterface  // Interfaces work like classes for context
		      
		    ElseIf IsClassDeclaration(line) Then
		      currentClass = ProcessClassDeclaration(line, currentModule, fileName, results)
		      currentInterface = ""
		      
		    ElseIf IsMethodOrFunctionDeclaration(line) Then
		      ProcessMethodDeclaration(line, currentModule, currentClass, fileName, results)
		      inMethodOrFunction = True
		      
		    ElseIf IsPropertyDeclaration(line) Then
		      ProcessPropertyDeclaration(line, currentModule, currentClass, fileName, results)
		      
		    ElseIf IsVariableDeclaration(line) And Not inMethodOrFunction Then
		      ProcessVariableDeclaration(line, currentModule, currentClass, fileName, results)
		      
		    ElseIf line.BeginsWith("#tag ComputedProperty") Or line.BeginsWith("#tag Event") Then
		      inMethodOrFunction = True
		      
		    ElseIf line.BeginsWith("#tag EndModule") Then
		      currentModule = ""
		      currentClass = ""
		      currentInterface = ""
		      inMethodOrFunction = False
		      
		    ElseIf line.BeginsWith("#tag EndClass") Then
		      currentClass = ""
		      currentInterface = ""
		      inMethodOrFunction = False
		      
		    ElseIf line.BeginsWith("#tag EndInterface") Then
		      currentClass = ""
		      currentInterface = ""
		      inMethodOrFunction = False
		      
		    ElseIf line.BeginsWith("#tag End") Or line = "End" Or line = "End Sub" Or line = "End Function" Then
		      inMethodOrFunction = False
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ProcessClassDeclaration(line As String, currentModule As String, fileName As String, ByRef results() As String) As String
		  // Private Function ProcessClassDeclaration(line As String, currentModule As String, fileName As String, ByRef results() As String) As String
		  Var className As String = line
		  
		  // Remove visibility modifiers
		  className = className.Replace("Protected Class ", "")
		  className = className.Replace("Public Class ", "")
		  className = className.Replace("Private Class ", "")
		  className = className.Replace("Class ", "")
		  
		  // Remove inheritance/implements clauses
		  Var spacePos As Integer = className.IndexOf(" ")
		  If spacePos > 0 Then
		    className = className.Left(spacePos)
		  End If
		  
		  className = className.Trim
		  
		  // Skip system classes
		  If className <> "" And className <> "App" Then
		    Var fullName As String
		    If currentModule <> "" Then
		      fullName = "CLASS: " + currentModule + "." + className
		    Else
		      fullName = "CLASS: " + className
		    End If
		    
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found class: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		  Return className
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessConstantDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  // rivate Sub ProcessConstantDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  Var parts() As String = line.Split("Name = ")
		  If parts.LastIndex >= 1 Then
		    Var namepart As String = parts(1)
		    Var commaPos As Integer = namepart.IndexOf(",")
		    If commaPos > 0 Then
		      Var constName As String = namepart.Left(commaPos).Trim
		      Var fullName As String = BuildFullName("CONSTANT", constName, currentModule, currentClass)
		      
		      If results.IndexOf(fullName) = -1 Then
		        results.Add(fullName)
		        System.DebugLog("Found constant: " + fullName + " in " + fileName)
		      End If
		    End If
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ProcessInterfaceDeclaration(line As String, currentModule As String, fileName As String, ByRef results() As String) As String
		  // Private Function ProcessInterfaceDeclaration(line As String, currentModule As String, fileName As String, ByRef results() As String) As String
		  Var interfaceName As String = line
		  
		  // Remove visibility modifiers
		  interfaceName = interfaceName.Replace("Protected Interface ", "")
		  interfaceName = interfaceName.Replace("Public Interface ", "")
		  interfaceName = interfaceName.Replace("Private Interface ", "")
		  interfaceName = interfaceName.Replace("Interface ", "")
		  
		  // Remove inheritance clauses
		  Var spacePos As Integer = interfaceName.IndexOf(" ")
		  If spacePos > 0 Then
		    interfaceName = interfaceName.Left(spacePos)
		  End If
		  
		  interfaceName = interfaceName.Trim
		  
		  If interfaceName <> "" Then
		    Var fullName As String
		    If currentModule <> "" Then
		      fullName = "INTERFACE: " + currentModule + "." + interfaceName
		    Else
		      fullName = "INTERFACE: " + interfaceName
		    End If
		    
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found interface: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		  Return interfaceName
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessMethodDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  // Private Sub ProcessMethodDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  Var methodName As String = ExtractMethodName(line)
		  
		  If methodName <> "" And Not IsSystemMethod(methodName) Then
		    Var fullName As String = BuildFullName("METHOD", methodName, currentModule, currentClass)
		    
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found method: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ProcessModuleDeclaration(line As String, fileName As String, ByRef results() As String) As String
		  Var moduleName As String = line
		  
		  // Remove visibility modifiers
		  moduleName = moduleName.Replace("Protected Module ", "")
		  moduleName = moduleName.Replace("Public Module ", "")
		  moduleName = moduleName.Replace("Private Module ", "")
		  moduleName = moduleName.Replace("Global Module ", "")
		  moduleName = moduleName.Replace("Module ", "")
		  
		  // Remove inheritance/implements clauses
		  Var spacePos As Integer = moduleName.IndexOf(" ")
		  If spacePos > 0 Then
		    moduleName = moduleName.Left(spacePos)
		  End If
		  
		  moduleName = moduleName.Trim
		  
		  If moduleName <> "" Then
		    Var fullName As String = "MODULE: " + moduleName
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found module: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		  Return moduleName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessPropertyDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  // Private Sub ProcessPropertyDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  Var propName As String = ExtractPropertyName(line)
		  
		  If propName <> "" Then
		    Var fullName As String = BuildFullName("PROPERTY", propName, currentModule, currentClass)
		    
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found property: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessSourceFile((item As FolderItem, ByRef results() As String))
		  
		  // Private Sub ProcessSourceFile(item As FolderItem, ByRef results() As String)
		  Try
		    // Read file content
		    Var tis As TextInputStream = TextInputStream.Open(item)
		    If tis = Nil Then Return
		    
		    Var content As String = tis.ReadAll
		    tis.Close
		    
		    // Parse the content
		    ParseFileContent(content, item.Name, results)
		    
		  Catch e As IOException
		    System.DebugLog("Error reading file: " + item.Name + " - " + e.Message)
		  End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ProcessVariableDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  // Private Sub ProcessVariableDeclaration(line As String, currentModule As String, currentClass As String, fileName As String, ByRef results() As String)
		  Var varName As String = ExtractVariableName(line)
		  
		  If varName <> "" Then
		    Var fullName As String = BuildFullName("VARIABLE", varName, currentModule, currentClass)
		    
		    If results.IndexOf(fullName) = -1 Then
		      results.Add(fullName)
		      System.DebugLog("Found variable: " + fullName + " in " + fileName)
		    End If
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScanProjectForDeclarations(folder As FolderItem, ByRef results() As String)
		  //Sub ScanProjectForDeclarations(folder As FolderItem, ByRef results() As String)
		  // Validate input
		  If folder = Nil Or Not folder.Exists Then Return
		  
		  // Process each item in the folder
		  For Each item As FolderItem In folder.Children
		    Try
		      If item = Nil Then Continue
		      
		      If item.IsFolder Then
		        // Recursively scan subfolders
		        ScanProjectForDeclarations(item, results)
		      ElseIf IsXojoSourceFile(item) Then
		        // Process Xojo source files
		        ProcessSourceFile(item, results)
		      End If
		      
		    Catch e As RuntimeException
		      System.DebugLog("Error processing item: " + If(item <> Nil, item.Name, "unknown") + " - " + e.Message)
		      Continue
		    End Try
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
		    ElseIf IsXojoSourceFile(item) Then
		      
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
		            // Extract the actual method name from the full path
		            Var methodParts() As String = word.Split(".")
		            Var actualMethodName As String = methodParts(methodParts.LastIndex)
		            
		            // Check for direct calls
		            If cleanedText.IndexOf(actualMethodName + "(") >= 0 Then
		              found = True
		              whereFound = "called in " + item.Name
		              // Check for qualified calls (Module.Method or Class.Method)
		            ElseIf cleanedText.IndexOf(word + "(") >= 0 Then
		              found = True
		              whereFound = "fully qualified call in " + item.Name
		              // AddressOf references
		            ElseIf cleanedText.IndexOf("AddressOf " + actualMethodName) >= 0 Or _
		              cleanedText.IndexOf("AddressOf " + word) >= 0 Then
		              found = True
		              whereFound = "referenced via AddressOf in " + item.Name
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
		            
		          Case "MODULE"
		            // Modules are used when their methods/properties are called
		            // Check for ModuleName.MethodName or ModuleName.PropertyName patterns
		            If cleanedText.IndexOf(word + ".") >= 0 Then
		              found = True
		              whereFound = "module referenced in " + item.Name
		            End If
		            // Also check if module is referenced in Extends or Implements
		            If cleanedText.IndexOf("Extends " + word) >= 0 Or _
		              cleanedText.IndexOf("Implements " + word) >= 0 Then
		              found = True
		              whereFound = "module extended/implemented in " + item.Name
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
		            
		          Case "INTERFACE"
		            // Interfaces are used when implemented
		            If cleanedText.IndexOf("Implements " + word) >= 0 Then
		              found = True
		              whereFound = "implemented in " + item.Name
		            End If
		            // Check if interface type is used in method signatures or properties
		            If cleanedText.IndexOf("As " + word) >= 0 Then
		              found = True  
		              whereFound = "used as type in " + item.Name
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
		  Var unusedModules() As String
		  Var unusedInterfaces() As String
		  
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
		      ElseIf item.BeginsWith("MODULE:") Then  // Add this
		        unusedModules.Add(item)
		      ElseIf item.BeginsWith("PROPERTY:") Then
		        unusedProperties.Add(item)
		      ElseIf item.BeginsWith("CONSTANT:") Then
		        unusedConstants.Add(item)
		      ElseIf item.BeginsWith("CLASS:") Then
		        unusedClasses.Add(item)
		      ElseIf item.BeginsWith("VARIABLE:") Then
		        unusedVariables.Add(item)
		      ElseIf item.BeginsWith("INTERFACE:") Then
		        unusedInterfaces.Add(item)
		      End If
		    End If
		  Next
		  
		  // Build categorized output
		  Var output As String = "FINAL ANALYSIS RESULTS" + EndOfLine
		  output = output + "Total items scanned: " + declaredItems.Count.ToString + EndOfLine + EndOfLine
		  
		  // Update the total count:
		  Var totalUnused As Integer = unusedMethods.Count + unusedProperties.Count + _
		  unusedConstants.Count + unusedClasses.Count + unusedVariables.Count + _
		  unusedModules.Count + unusedInterfaces.Count
		  
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
		    
		    If unusedModules.Count > 0 Then
		      output = output + "UNUSED MODULES (" + unusedModules.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedModules
		        output = output + "  • " + item.Replace("MODULE: ", "") + EndOfLine
		      Next
		      output = output + EndOfLine
		    End If
		    
		    If unusedInterfaces.Count > 0 Then
		      output = output + "UNUSED INTERFACES (" + unusedInterfaces.Count.ToString + "):" + EndOfLine
		      For Each item As String In unusedInterfaces
		        output = output + "  • " + item.Replace("INTERFACE: ", "") + EndOfLine
		      Next
		      // output = output + EndOfLine
		    End If
		    
		    
		  End If
		  
		  
		  // ============ ADD THE SUMMARY SECTION HERE ============
		  output = output + EndOfLine + EndOfLine  // Extra spacing
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + "SUMMARY BY MODULE/CLASS:" + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine + EndOfLine
		  
		  // Create a dictionary to count unused items per module/class
		  Var contextCounts As New Dictionary
		  Var standaloneCount As Integer = 0
		  
		  
		  // Process all unused items to count by context
		  Var allUnusedItems() As String
		  
		  // Add all items from each array
		  For Each item As String In unusedMethods
		    allUnusedItems.Add(item)
		  Next
		  For Each item As String In unusedProperties
		    allUnusedItems.Add(item)
		  Next
		  For Each item As String In unusedConstants
		    allUnusedItems.Add(item)
		  Next
		  For Each item As String In unusedVariables
		    allUnusedItems.Add(item)
		  Next
		  For Each item As String In unusedClasses
		    allUnusedItems.Add(item)
		  Next
		  For Each item As String In unusedInterfaces
		    allUnusedItems.Add(item)
		  Next
		  
		  For Each item As String In allUnusedItems
		    // Remove the type prefix (METHOD:, PROPERTY:, etc.)
		    Var colonPos As Integer = item.IndexOf(": ")
		    If colonPos > 0 Then
		      Var itemPath As String = item.Middle(colonPos + 2)  // Get everything after ": "
		      Var parts() As String = itemPath.Split(".")
		      
		      If parts.Count > 1 Then
		        // Has module or class context
		        Var contextName As String = parts(0)  // First part is module or class
		        If parts.Count > 2 Then
		          // Has both module and class (module.class.item)
		          contextName = parts(0) + "." + parts(1)
		        End If
		        
		        If contextCounts.HasKey(contextName) Then
		          contextCounts.Value(contextName) = contextCounts.Value(contextName) + 1
		        Else
		          contextCounts.Value(contextName) = 1
		        End If
		      Else
		        // Standalone item (no module/class context)
		        standaloneCount = standaloneCount + 1
		      End If
		    End If
		  Next
		  
		  // Sort the keys for better readability
		  Var sortedKeys() As String
		  For Each key As String In contextCounts.Keys
		    sortedKeys.Add(key)
		  Next
		  sortedKeys.Sort
		  
		  // Output the summary
		  For Each key As String In sortedKeys
		    Var count As Integer = contextCounts.Value(key)
		    Var itemText As String = If(count = 1, "item", "items")
		    output = output + "  📁 " + key + ": " + count.ToString + " unused " + itemText + EndOfLine
		  Next
		  For Each item As String In unusedModules
		    allUnusedItems.Add(item)
		  Next
		  
		  If standaloneCount > 0 Then
		    Var itemText As String = If(standaloneCount = 1, "item", "items")
		    output = output + "  📄 [Global/Standalone]: " + standaloneCount.ToString + " unused " + itemText + EndOfLine
		  End If
		  
		  output = output + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  output = output + "Total unused items: " + totalUnused.ToString + EndOfLine
		  output = output + "═══════════════════════════════════════" + EndOfLine
		  // ============ END OF SUMMARY SECTION ============
		  
		  
		  // Finally, display everything
		  txtResults.Text = txtResults.Text + output
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExportButton
	#tag Event
		Sub Pressed()
		  Var saveFile As FolderItem = FolderItem.ShowSaveFileDialog("", "UnusedCode_" + DateTime.Now.ToString + ".txt")
		  If saveFile <> Nil Then
		    Try
		      Var tos As TextOutputStream = TextOutputStream.Create(saveFile)
		      tos.Write(txtResults.Text)
		      tos.Close
		      MessageBox("Results exported successfully!")
		    Catch e As IOException
		      MessageBox("Error saving file: " + e.Message)
		    End Try
		  End If
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
