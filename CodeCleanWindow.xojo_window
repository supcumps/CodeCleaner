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
   Title           =   "Code Analyzer - Scan and Visualize"
   Type            =   0
   Visible         =   True
   Width           =   828
   Begin DesktopBevelButton btnScan
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Scan .Xojo_Project File Folder"
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
      Left            =   43
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
      Width           =   194
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
      Height          =   530
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
      Top             =   125
      Transparent     =   False
      Underline       =   False
      UnicodeMode     =   1
      ValidationMask  =   ""
      Visible         =   True
      Width           =   743
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
      Left            =   722
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
      Top             =   13
      Transparent     =   False
      Visible         =   True
      Width           =   64
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopBevelButton ExportButton
      Active          =   False
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
      Enabled         =   False
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
      Left            =   43
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
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
      _mIndex         =   0
      _mInitialParent =   ""
      _mName          =   ""
      _mPanelIndex    =   0
   End
   Begin DesktopBevelButton btnGenerateFlowchart
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Generate Flowchart PDF"
      CaptionAlignment=   3
      CaptionDelta    =   0
      CaptionPosition =   1
      Enabled         =   False
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
      Left            =   531
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
      Scope           =   2
      TabIndex        =   4
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
   Begin DesktopPopupMenu popLayoutType
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      InitialValue    =   "Simple Layout\nCompact Layout\nHierarchical Layout"
      Italic          =   False
      Left            =   531
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      SelectedRowIndex=   0
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   88
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   179
   End
   Begin DesktopCheckBox chkShowRelationships
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Show Relationships"
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   531
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   54
      Transparent     =   False
      Underline       =   False
      Value           =   True
      Visible         =   True
      VisualState     =   0
      Width           =   153
   End
   Begin DesktopBevelButton btnRefactoringSuggestions
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Generate Refactoring Report"
      CaptionAlignment=   3
      CaptionDelta    =   0
      CaptionPosition =   1
      Enabled         =   False
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
      Left            =   234
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
      Scope           =   2
      TabIndex        =   7
      TabPanelIndex   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   54
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
   Begin DesktopBevelButton GenerateHotSpotsPDFButton
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Generate Hot Spots Report"
      CaptionAlignment=   3
      CaptionDelta    =   0
      CaptionPosition =   1
      Enabled         =   False
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
      Left            =   43
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
      Scope           =   2
      TabIndex        =   8
      TabPanelIndex   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   88
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
   Begin DesktopBevelButton btnShowAllHotSpots
      Active          =   False
      AllowAutoDeactivate=   True
      AllowFocus      =   True
      AllowTabStop    =   True
      BackgroundColor =   &c00000000
      BevelStyle      =   0
      Bold            =   False
      ButtonStyle     =   0
      Caption         =   "Show All  Hot Spots "
      CaptionAlignment=   3
      CaptionDelta    =   0
      CaptionPosition =   1
      Enabled         =   False
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
      Left            =   234
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuStyle       =   0
      PanelIndex      =   0
      Scope           =   2
      TabIndex        =   9
      TabPanelIndex   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   88
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
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Property, Flags = &h0
		mAnalyzer As ProjectAnalyzer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLastScannedFolder As FolderItem
	#tag EndProperty


#tag EndWindowCode

#tag Events btnScan
	#tag Event
		Sub Pressed()
		  // Select the Xojo project folder
		  Var folder As FolderItem = FolderItem.ShowSelectFolderDialog
		  
		  If folder = Nil Or Not folder.Exists Then
		    Return
		  End If
		  
		  // Store the folder for later use
		  mLastScannedFolder = folder
		  
		  // Clear previous results
		  txtResults.Text = ""
		  txtResults.Text = "Scanning project files..." + EndOfLine + EndOfLine
		  
		  // Create new analyzer
		  mAnalyzer = New ProjectAnalyzer
		  
		  // Scan the project
		  mAnalyzer.ScanProject(folder)
		  
		  // After mAnalyzer.ScanProject(folder)
		  System.DebugLog("=== CHECKING ELEMENTS ===")
		  
		  Var all() As CodeElement = mAnalyzer.GetAllElements()
		  System.DebugLog("GetAllElements: " + all.Count.ToString)
		  
		  If all.Count > 0 Then
		    System.DebugLog("First element: " + all(0).Name + " (Type: " + all(0).ElementType + ")")
		    System.DebugLog("Second element: " + all(1).Name + " (Type: " + all(1).ElementType + ")")
		  End If
		  
		  // Build relationships between elements
		  mAnalyzer.BuildRelationships(folder)
		  
		  System.DebugLog("=== ELEMENT TYPE CHECK ===")
		  
		  all()  = mAnalyzer.GetAllElements()
		  Var classes() As CodeElement = mAnalyzer.GetClassElements()
		  Var modules() As CodeElement = mAnalyzer.GetModuleElements()
		  Var methods() As CodeElement = mAnalyzer.GetMethodElements()
		  
		  Var methodsWithCode As Integer = 0
		  For Each m As CodeElement In methods
		    If m.Code.Trim <> "" Then
		      methodsWithCode = methodsWithCode + 1
		    End If
		  Next
		  System.DebugLog("Methods with code: " + methodsWithCode.ToString + " of " + methods.Count.ToString)
		  System.DebugLog("Methods with calls: " + methods.Count.ToString)  // Will show how many have CallsTo populated
		  
		  
		  System.DebugLog("Total: " + all.Count.ToString)
		  System.DebugLog("Classes: " + classes.Count.ToString)
		  System.DebugLog("Modules: " + modules.Count.ToString)
		  System.DebugLog("Methods: " + methods.Count.ToString)
		  
		  // Show first few elements and their types
		  For i As Integer = 0 To Min(5, all.Count - 1)
		    System.DebugLog("Element " + i.ToString + ": " + all(i).Name + " (Type: '" + all(i).ElementType + "')")
		  Next
		  
		  // Analyze error handling patterns
		  mAnalyzer.AnalyzeErrorHandling()
		  // DEBUG: Test parameter parsing
		  System.DebugLog("=== PARAMETER PARSING DEBUG ===")
		  methods() = mAnalyzer.GetMethodElements
		  Var testCount As Integer = 0
		  
		  For Each method As CodeElement In methods
		    If method.Code.Trim <> "" And testCount < 5 Then  // Test first 5 methods
		      System.DebugLog("")
		      System.DebugLog("Method: " + method.FullPath)
		      System.DebugLog("Code starts with:")
		      
		      // Show first 3 lines of code
		      Var lines() As String = method.Code.Split(EndOfLine)
		      For i As Integer = 0 To Min(2, lines.Count - 1)
		        System.DebugLog("  Line " + i.ToString + ": " + lines(i))
		      Next
		      
		      // Test parsing
		      Var result As Dictionary = mAnalyzer.ParseMethodParameters(method.Code)
		      Var paramCount As Integer = result.Value("parameterCount")
		      System.DebugLog("Detected parameters: " + paramCount.ToString)
		      
		      testCount = testCount + 1
		    End If
		  Next
		  System.DebugLog("=== END DEBUG ===")
		  System.DebugLog("=== PARAMETER EXTRACTION TEST ===")
		  
		  Var methodsWithParams As Integer = 0
		  Var totalParams As Integer = 0
		  
		  For Each method As CodeElement In methods
		    If method.ParameterCount > 0 Then
		      methodsWithParams = methodsWithParams + 1
		      totalParams = totalParams + method.ParameterCount
		      System.DebugLog("✓ " + method.Name + ": " + method.ParameterCount.ToString + " params")
		      System.DebugLog("  Parameters: " + method.Parameters)
		    End If
		  Next
		  
		  System.DebugLog("")
		  System.DebugLog("Methods with parameters: " + methodsWithParams.ToString)
		  System.DebugLog("Total parameters: " + totalParams.ToString)
		  
		  // TEST: Check if code is being captured
		  methods() = manalyzer.GetMethodElements
		  System.DebugLog("=== CODE CAPTURE TEST ===")
		  System.DebugLog("Total methods found: " + methods.Count.ToString)
		  
		  methodsWithCode = 0
		  For Each method As CodeElement In methods
		    If method.Code.Trim <> "" Then
		      methodsWithCode = methodsWithCode + 1
		      System.DebugLog("✅ " + method.Name + " has " + method.Code.Length.ToString + " chars")
		    End If
		  Next
		  
		  System.DebugLog("Methods with code: " + methodsWithCode.ToString + " of " + methods.Count.ToString)
		  
		  If methodsWithCode = 0 Then
		    System.DebugLog("❌ NO CODE CAPTURED! Need to update CodeElement and ProjectAnalyzer")
		  Else
		    System.DebugLog("✅ Code is being captured!")
		  End If
		  
		  // Generate text report (keeping your existing functionality)
		  GenerateTextReport()
		  
		  // Enable the export and flowchart buttons
		  ExportButton.Enabled = True
		  btnGenerateFlowchart.Enabled = True
		  btnRefactoringSuggestions.Enabled = True
		  GenerateHotSpotsPDFButton.Enabled = True
		  btnShowAllHotSpots.Enabled =True
		  
		  Var allElements() As CodeElement = mAnalyzer.GetAllElements()
		  MessageBox("Scan complete! Found " + allElements.Count.ToString + " code elements.")
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ExportButton
	#tag Event
		Sub Pressed()
		  
		  If mAnalyzer = Nil Then
		    MessageBox("Please scan a project first!")
		    Return
		  End If
		  
		  Var dt As DateTime = DateTime.Now
		  Var timestamp As String = dt.Year.ToString + "-" + _
		  Format(dt.Month, "00") + "-" + _
		  Format(dt.Day, "00") + "_" + _
		  Format(dt.Hour, "00") + Format(dt.Minute, "00")
		  
		  Var saveFile As FolderItem = FolderItem.ShowSaveFileDialog("application/pdf", "CodeAnalysis_" + timestamp + ".pdf")
		  If saveFile <> Nil Then
		    Try
		      Var generator As New ReportGenerator
		      If generator.GenerateAnalysisReportPDF(mAnalyzer, saveFile) Then
		        MessageBox("PDF report generated successfully!")
		      Else
		        MessageBox("Error generating PDF report. Check the console for details.")
		      End If
		    Catch e As RuntimeException
		      MessageBox("Error saving PDF: " + e.Message)
		    End Try
		  End If
		  
		  
		  
		  
		  
		  'Var timestamp As String = dt.Year.ToString + "-" + _
		  'Format(dt.Month, "00") + "-" + _
		  'Format(dt.Day, "00") + "_" + _
		  'Format(dt.Hour, "00") + Format(dt.Minute, "00")
		  '
		  'Var saveFile As FolderItem = FolderItem.ShowSaveFileDialog("", "UnusedCode_" + timestamp + ".txt")
		  'If saveFile <> Nil Then
		  'Try
		  'Var tos As TextOutputStream = TextOutputStream.Create(saveFile)
		  'tos.Write(txtResults.Text)
		  'tos.Close
		  'MessageBox("Results exported successfully!")
		  'Catch e As IOException
		  'MessageBox("Error saving file: " + e.Message)
		  'End Try
		  'End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnGenerateFlowchart
	#tag Event
		Sub Pressed()
		  If mAnalyzer = Nil Then
		    MessageBox("Please scan a project first!")
		    Return
		  End If
		  
		  // Get selected layout type
		  Var layoutType As String = "Simple"
		  If popLayoutType.SelectedRowIndex = 1 Then
		    layoutType = "Compact"
		  ElseIf popLayoutType.SelectedRowIndex = 2 Then
		    layoutType = "Hierarchical"
		  End If
		  
		  // Get relationship option
		  Var includeRelationships As Boolean = chkShowRelationships.Value
		  
		  // Ask user what to visualize
		  Var dlg As New MessageDialog
		  dlg.Title = "Select Elements to Visualize"
		  dlg.Message = "What would you like to include in the flowchart?"
		  dlg.ActionButton.Caption = "All Elements"
		  dlg.AlternateActionButton.Caption = "Unused Only"
		  dlg.CancelButton.Caption = "Cancel"
		  
		  Var result As MessageDialogButton = dlg.ShowModal
		  
		  If result = dlg.CancelButton Then
		    Return
		  End If
		  
		  // Get elements to visualize
		  Var elementsToShow() As CodeElement
		  If result = dlg.AlternateActionButton Then
		    elementsToShow = mAnalyzer.GetUnusedElements()
		    If elementsToShow.Count = 0 Then
		      MessageBox("No unused elements found!")
		      Return
		    End If
		  Else
		    elementsToShow = mAnalyzer.AllElements
		  End If
		  
		  // Get save location
		  Var dt As DateTime = DateTime.Now
		  Var timestamp As String = dt.Year.ToString + "-" + _
		  Format(dt.Month, "00") + "-" + _
		  Format(dt.Day, "00") + "_" + _
		  Format(dt.Hour, "00") + Format(dt.Minute, "00")
		  
		  Var saveFile As FolderItem = FolderItem.ShowSaveFileDialog("application/pdf", "Flowchart_" + timestamp + ".pdf")
		  If saveFile = Nil Then
		    Return
		  End If
		  
		  // Generate the flowchart
		  Var generator As New FlowchartGenerator
		  
		  txtResults.Text = txtResults.Text + EndOfLine + "Generating flowchart PDF..." + EndOfLine
		  
		  If generator.GenerateFlowchartPDF(elementsToShow, saveFile, layoutType, includeRelationships) Then
		    MessageBox("Flowchart PDF generated successfully!")
		  Else
		    MessageBox("Error generating flowchart PDF. Check the console for details.")
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnRefactoringSuggestions
	#tag Event
		Sub Pressed()
		  If mAnalyzer = Nil Then
		    MessageBox("Please scan a project first!")
		    Return
		  End If
		  
		  // Check if we have suggestions
		  Var allMethods() As CodeElement = mAnalyzer.GetMethodElements()
		  Var hasSuggestions As Boolean = False
		  
		  For Each method As CodeElement In allMethods
		    If method.RefactoringSuggestions.Count > 0 Then
		      hasSuggestions = True
		      Exit For
		    End If
		  Next
		  
		  If Not hasSuggestions Then
		    MessageBox("No refactoring suggestions found. Your code looks good!")
		    Return
		  End If
		  
		  // Get save location
		  Var dlg As New SaveFileDialog
		  dlg.SuggestedFileName = "RefactoringSuggestions_" + DateTime.Now.ToString("yyyy-MM-dd_HHmm") + ".pdf"
		  dlg.Filter = "PDF Files (*.pdf)|*.pdf"
		  
		  Var saveFile As FolderItem = dlg.ShowModal()
		  If saveFile = Nil Then Return
		  
		  // Generate the PDF
		  Var generator As New ReportGenerator
		  If generator.GenerateRefactoringSuggestionsReport(mAnalyzer, saveFile) Then
		    MessageBox("Refactoring suggestions report generated successfully!")
		    saveFile.Open()
		  Else
		    MessageBox("Error generating refactoring suggestions report.")
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events GenerateHotSpotsPDFButton
	#tag Event
		Sub Pressed()
		  // Button: GenerateHotSpotsPDFButton
		  
		  // Sub GenerateHotSpotsPDFButton_Pressed()
		  If mAnalyzer = Nil Then
		    MessageBox("Please scan a project first")
		    Return
		  End If
		  
		  // Get method elements - use GetMethodElements() like the working button
		  Var elements() As CodeElement = mAnalyzer.GetMethodElements()
		  
		  If elements.Count = 0 Then
		    MessageBox("No methods found in the scanned project")
		    Return
		  End If
		  
		  // Generate hot spots
		  Var hotSpots() As HotSpot = HotSpotsGenerator.GenerateHotSpots(elements)
		  
		  If hotSpots.Count = 0 Then
		    MessageBox("No significant hot spots detected in this project!")
		    Return
		  End If
		  
		  // Ask user where to save
		  Var dlg As New SaveFileDialog
		  dlg.SuggestedFileName = "HotSpots_Report.pdf"
		  dlg.Filter = "PDF Files|*.pdf"
		  
		  Var f As FolderItem = dlg.ShowModal
		  If f <> Nil Then
		    Var generator As New ReportGenerator
		    
		    // Convert FolderItem to String
		    Var projectName As String
		    If mLastScannedFolder <> Nil Then
		      projectName = mLastScannedFolder.Name
		    Else
		      projectName = "Unknown Project"
		    End If
		    
		    generator.GenerateHotSpotsReportPDF(hotSpots, f.NativePath, projectName)
		    
		    MessageBox("Hot Spots report generated successfully!" + EndOfLine + _
		    "Found " + hotSpots.Count.ToString + " hot spots")
		  End If
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnShowAllHotSpots
	#tag Event
		Sub Pressed()
		  // Button: btnShowAllHotSpots
		  
		  // Sub btnShowAllHotSpots_Pressed()
		  If mAnalyzer = Nil Then
		    MessageBox("Please scan a project first")
		    Return
		  End If
		  
		  // Get method elements - use GetMethodElements() like the working button
		  Var elements() As CodeElement = mAnalyzer.GetMethodElements()
		  
		  If elements.Count = 0 Then
		    MessageBox("No methods found in the scanned project")
		    Return
		  End If
		  
		  // Generate hot spots
		  Var hotSpots() As HotSpot = HotSpotsGenerator.GenerateHotSpots(elements)
		  
		  If hotSpots.Count = 0 Then
		    txtResults.Text = "✓ No significant hot spots detected in this project!"
		    Return
		  End If
		  
		  // Generate and display text report
		  Var report As String = CodeCleanWindowHelpers.GenerateHotSpotsTextReport(hotSpots)
		  txtResults.Text = report
		  
		  // Optional: Scroll to top
		  txtResults.SelectionStart = 0
		  txtResults.SelectionLength = 0
		  
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
