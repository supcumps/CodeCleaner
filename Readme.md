# CodeAnalysis

## Project Overview

This Xojo project contains the following components:

## Windows

### CodeCleanWindow

#### Properties

- **DefaultLocation**: 2
- **HasCloseButton**: True
- **HasMaximizeButton**: True
- **HasMinimizeButton**: True
- **Height**: 20
- **MaximumHeight**: 32000
- **MaximumWidth**: 32000
- **MinimumHeight**: 64
- **MinimumWidth**: 64
- **Resizeable**: True
- **Title**: "Code Analyzer - Scan and Visualize"
- **Type**: "Boolean"
- **Visible**: true
- **Width**: 153
- **Backdrop**: 0
- **BackgroundColor**: &c00000000
- **Composite**: False
- **FullScreen**: False
- **HasBackgroundColor**: False
- **HasFullScreenButton**: False
- **HasTitleBar**: True
- **ImplicitInstance**: True
- **MacProcID**: 0
- **MenuBar**: 257478655
- **MenuBarVisible**: False
- **Active**: False
- **AllowAutoDeactivate**: True
- **AllowFocus**: True
- **AllowTabStop**: True
- **BevelStyle**: 0
- **Bold**: False
- **ButtonStyle**: 0
- **Caption**: "Show Relationships"
- **CaptionAlignment**: 3
- **CaptionDelta**: 0
- **CaptionPosition**: 1
- **Enabled**: True
- **FontName**: "System"
- **FontSize**: 0.0
- **FontUnit**: 0
- **Icon**: 0
- **IconAlignment**: 0
- **IconDeltaX**: 0
- **IconDeltaY**: 0
- **Index**: -2147483648
- **InitialParent**: ""
- **Italic**: False
- **Left**: 440
- **LockBottom**: False
- **LockedInPosition**: False
- **LockLeft**: True
- **LockRight**: False
- **LockTop**: True
- **MenuStyle**: 0
- **PanelIndex**: 0
- **Scope**: 2
- **TabIndex**: 6
- **TabPanelIndex**: 0
- **TextColor**: &c00000000
- **Tooltip**: ""
- **Top**: 23
- **Transparent**: False
- **Underline**: False
- **Value**: True
- **AllowFocusRing**: True
- **AllowSpellChecking**: True
- **AllowStyledText**: True
- **AllowTabs**: False
- **Format**: ""
- **HasBorder**: True
- **HasHorizontalScrollbar**: False
- **HasVerticalScrollbar**: True
- **HideSelection**: True
- **LineHeight**: 0.0
- **LineSpacing**: 1.0
- **MaximumCharactersAllowed**: 0
- **Multiline**: True
- **ReadOnly**: True
- **TabStop**: True
- **Text**: ""
- **TextAlignment**: 0
- **UnicodeMode**: 1
- **ValidationMask**: ""
- **Image**: 1280350207
- **InitialValue**: "False"
- **SelectedRowIndex**: 0
- **VisualState**: 0
- **Name**: "MenuBarVisible"
- **Group**: "Deprecated"
- **EditorType**: ""

#### Events

##### btnScan:
- **Sub Pressed()**

##### ExportButton:
- **Sub Pressed()**

##### btnGenerateFlowchart:
- **Sub Pressed()**

---

## Project Components

- **Classes:** 6 (App, ProjectAnalyzer, CodeElement, FlowchartGenerator, ReportGenerator, MethodMetrics)
- **Modules:** 1 (CodeCleanWindowHelpers)
- **Windows:** 1 (CodeCleanWindow)
- **Menus:** 1 (MainMenuBar)

## Classes

### App

#### Properties

- **`kEditClear`** Public String

- **`kFileQuit`** Public String

- **`kFileQuitShortcut`** Public String

#### Methods

None

#### Events

None

---

### ProjectAnalyzer

**Type:** Interface

#### Properties

- **`AllElements`** Public CodeElement()

- **`ClassElements`** Public CodeElement()

- **`ElementLookup`** Public Dictionary

- **`MethodElements`** Public CodeElement()

- **`ModuleElements`** Public CodeElement()

#### Methods

- **`AnalyzeFileForRelationships`** Private Sub
  - **Parameters:** `content As String`
  - **Signature:** `Private Sub AnalyzeFileForRelationships(content As String)`

- **`BuildFullName`** Private Function
  - **Parameters:** `itemType As String, itemName As String, currentModule As String, currentClass As String`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildFullName(itemType As String, itemName As String, currentModule As String, currentClass As String) As String`

- **`BuildMethodFullPath`** Private Function
  - **Parameters:** `line As String, currentModule As String, currentClass As String`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildMethodFullPath(line As String, currentModule As String, currentClass As String) As String`

- **`BuildRelationships`** Public Sub
  - **Parameters:** `projectFolder As FolderItem`
  - **Signature:** `Public Sub BuildRelationships(projectFolder As FolderItem)`

- **`CheckForDuplicates`** Public Function
  - **Returns:** `String`
  - **Signature:** `Public Function CheckForDuplicates() As String`

- **`CleanCodeForAnalysis`** Private Function
  - **Parameters:** `content As String`
  - **Returns:** `String`
  - **Signature:** `Private Function CleanCodeForAnalysis(content As String) As String`

- **`CollectCallChain`** Private Sub
  - **Parameters:** `element As CodeElement, ByRef chain() As CodeElement, visited As Dictionary, currentDepth As Integer, maxDepth As Integer`
  - **Signature:** `Private Sub CollectCallChain(element As CodeElement, ByRef chain() As CodeElement, visited As Dictionary, currentDepth As Integer, maxDepth As Integer)`

- **`Constructor`** Public Constructor
  - **Signature:** `Public Constructor()`

- **`DetectMethodCalls`** Private Sub
  - **Parameters:** `line As String, callingMethodFullPath As String`
  - **Signature:** `Private Sub DetectMethodCalls(line As String, callingMethodFullPath As String)`

- **`ExtractClassName`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ExtractClassName(line As String) As String`

- **`ExtractMethodName`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ExtractMethodName(line As String) As String`

- **`ExtractModuleName`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ExtractModuleName(line As String) As String`

- **`ExtractPropertyName`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ExtractPropertyName(line As String) As String`

- **`ExtractVariableName`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ExtractVariableName(line As String) As String`

- **`FindElementByFullPath`** Public Function
  - **Parameters:** `fullPath As String`
  - **Returns:** `CodeElement`
  - **Signature:** `Public Function FindElementByFullPath(fullPath As String) As CodeElement`

- **`GetCallChain`** Public Function
  - **Parameters:** `startElement As CodeElement, maxDepth As Integer = 5`
  - **Returns:** `CodeElement()`
  - **Signature:** `Public Function GetCallChain(startElement As CodeElement, maxDepth As Integer = 5) As CodeElement()`

- **`GetClassElements`** Public Function
  - **Returns:** `CodeElement()`
  - **Signature:** `Public Function GetClassElements() As CodeElement()`

- **`GetMethodElements`** Public Function
  - **Returns:** `CodeElement()`
  - **Signature:** `Public Function GetMethodElements() As CodeElement()`

- **`GetModuleElements`** Public Function
  - **Returns:** `CodeElement()`
  - **Signature:** `Public Function GetModuleElements() As CodeElement()`

- **`GetUnusedElements`** Public Function
  - **Returns:** `CodeElement()`
  - **Signature:** `Public Function GetUnusedElements() As CodeElement()`

- **`HandleEndTag`** Private Function
  - **Parameters:** `line As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentInterface As String`
  - **Signature:** `Private Function HandleEndTag(line As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentInterface As String)`

- **`IsClassDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsClassDeclaration(line As String) As Boolean`

- **`IsEndTag`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsEndTag(line As String) As Boolean`

- **`IsInterfaceDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsInterfaceDeclaration(line As String) As Boolean`

- **`IsMethodOrFunctionDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsMethodOrFunctionDeclaration(line As String) As Boolean`

- **`IsModuleDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsModuleDeclaration(line As String) As Boolean`

- **`IsPropertyDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsPropertyDeclaration(line As String) As Boolean`

- **`IsSystemMethod`** Private Function
  - **Parameters:** `methodName As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsSystemMethod(methodName As String) As Boolean`

- **`IsVariableDeclaration`** Private Function
  - **Parameters:** `line As String`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsVariableDeclaration(line As String) As Boolean`

- **`IsXojoSourceFile`** Private Function
  - **Parameters:** `item As FolderItem`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsXojoSourceFile(item As FolderItem) As Boolean`

- **`MarkUsedElements`** Private Sub
  - **Parameters:** `cleanedText As String`
  - **Signature:** `Private Sub MarkUsedElements(cleanedText As String)`

- **`ParseFileContent`** Private Sub
  - **Parameters:** `content As String, fileName As String`
  - **Signature:** `Private Sub ParseFileContent(content As String, fileName As String)`

- **`ProcessClassDeclaration`** Private Function
  - **Parameters:** `line As String, currentModule As String, fileName As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ProcessClassDeclaration(line As String, currentModule As String, fileName As String) As String`

- **`ProcessInterfaceDeclaration`** Private Function
  - **Parameters:** `line As String, currentModule As String, fileName As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ProcessInterfaceDeclaration(line As String, currentModule As String, fileName As String) As String`

- **`ProcessLine`** Private Function
  - **Parameters:** `trimmedLine As String, fullLine As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentInterface As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String, fileName As String`
  - **Signature:** `Private Function ProcessLine(trimmedLine As String, fullLine As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentInterface As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String, fileName As String)`

- **`ProcessLineForRelationships`** Private Sub
  - **Parameters:** `line As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentMethodFullPath As String, ByRef inMethod As Boolean`
  - **Signature:** `Private Sub ProcessLineForRelationships(line As String, ByRef currentModule As String, ByRef currentClass As String, ByRef currentMethodFullPath As String, ByRef inMethod As Boolean)`

- **`ProcessMethodDeclaration`** Private Function
  - **Parameters:** `line As String, currentModule As String, currentClass As String, fileName As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ProcessMethodDeclaration(line As String, currentModule As String, currentClass As String, fileName As String) As String`

- **`ProcessModuleDeclaration`** Private Function
  - **Parameters:** `line As String, fileName As String`
  - **Returns:** `String`
  - **Signature:** `Private Function ProcessModuleDeclaration(line As String, fileName As String) As String`

- **`ProcessPropertyDeclaration`** Private Sub
  - **Parameters:** `line As String, currentModule As String, currentClass As String, fileName As String`
  - **Signature:** `Private Sub ProcessPropertyDeclaration(line As String, currentModule As String, currentClass As String, fileName As String)`

- **`ProcessSourceFile`** Private Sub
  - **Parameters:** `item As FolderItem`
  - **Signature:** `Private Sub ProcessSourceFile(item As FolderItem)`

- **`ProcessVariableDeclaration`** Private Sub
  - **Parameters:** `line As String, currentModule As String, currentClass As String, fileName As String`
  - **Signature:** `Private Sub ProcessVariableDeclaration(line As String, currentModule As String, currentClass As String, fileName As String)`

- **`ResetMethodContext`** Private Function
  - **Parameters:** `ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String`
  - **Signature:** `Private Function ResetMethodContext(ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String)`

- **`ScanFileForRelationships`** Private Sub
  - **Parameters:** `item As FolderItem`
  - **Signature:** `Private Sub ScanFileForRelationships(item As FolderItem)`

- **`ScanForRelationships`** Public Sub
  - **Parameters:** `folder As FolderItem`
  - **Signature:** `Public Sub ScanForRelationships(folder As FolderItem)`

- **`ScanProject`** Public Sub
  - **Parameters:** `folder As FolderItem`
  - **Signature:** `Public Sub ScanProject(folder As FolderItem)`

- **`ScanProjectForDeclarations`** Private Sub
  - **Parameters:** `folder As FolderItem`
  - **Signature:** `Private Sub ScanProjectForDeclarations(folder As FolderItem)`

- **`StartMethodCapture`** Private Function
  - **Parameters:** `line As String, currentModule As String, currentClass As String, fileName As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String`
  - **Signature:** `Private Function StartMethodCapture(line As String, currentModule As String, currentClass As String, fileName As String, ByRef inMethodOrFunction As Boolean, ByRef currentMethodFullPath As String, ByRef currentMethodCode As String)`

#### Events

None

---

### CodeElement

#### Properties

- **`CalledBy`** Public CodeElement()

- **`CallsTo`** Public CodeElement()

- **`Code`** Public String

- **`Contains`** Public CodeElement()

- **`ElementType`** Public String

- **`FileName`** Public String

- **`FullPath`** Public String

- **`Height`** Public Integer = 60

- **`IsUsed`** Public Boolean = False

- **`LayoutLevel`** Public Integer = 0

- **`ModuleName`** Public String

- **`Name`** Public String

- **`Parent`** Public CodeElement

- **`ParentClass`** Public String

- **`Width`** Public Integer = 180

- **`X`** Public Integer

- **`Y`** Public Integer

#### Methods

- **`Constructor`** Public Constructor
  - **Parameters:** `elementType As String, name As String, fullPath As String, Optional fileName As String = ""`
  - **Signature:** `Public Constructor(elementType As String, name As String, fullPath As String, Optional fileName As String = "")`

#### Events

None

---

### FlowchartGenerator

#### Properties

- **`HorizontalSpacing`** Public Integer = 50

- **`PageHeight`** Public Integer = 595

- **`PageMargin`** Public Integer = 30

- **`PageWidth`** Public Integer = 842

- **`VerticalSpacing`** Public Integer = 80

#### Methods

- **`CalculateCompactLayout`** Public Sub
  - **Parameters:** `elements() As CodeElement`
  - **Signature:** `Public Sub CalculateCompactLayout(elements() As CodeElement)`

- **`CalculateHierarchicalLayout`** Public Sub
  - **Parameters:** `elements() As CodeElement`
  - **Signature:** `Public Sub CalculateHierarchicalLayout(elements() As CodeElement)`

- **`CalculateSimpleLayout`** Public Sub
  - **Parameters:** `elements() As CodeElement`
  - **Signature:** `Public Sub CalculateSimpleLayout(elements() As CodeElement)`

- **`Constructor`** Public Constructor
  - **Signature:** `Public Constructor()`

- **`DrawArrow`** Private Sub
  - **Parameters:** `g As Graphics, fromElement As CodeElement, toElement As CodeElement`
  - **Signature:** `Private Sub DrawArrow(g As Graphics, fromElement As CodeElement, toElement As CodeElement)`

- **`DrawArrowWithOffset`** Private Sub
  - **Parameters:** `g As Graphics, fromElement As CodeElement, toElement As CodeElement, pageYOffset As Integer`
  - **Signature:** `Private Sub DrawArrowWithOffset(g As Graphics, fromElement As CodeElement, toElement As CodeElement, pageYOffset As Integer)`

- **`DrawConnections`** Private Sub
  - **Parameters:** `g As Graphics, elements() As CodeElement`
  - **Signature:** `Private Sub DrawConnections(g As Graphics, elements() As CodeElement)`

- **`DrawLegend`** Private Sub
  - **Parameters:** `g As Graphics`
  - **Signature:** `Private Sub DrawLegend(g As Graphics)`

- **`DrawNode`** Private Sub
  - **Parameters:** `g As Graphics, element As CodeElement`
  - **Signature:** `Private Sub DrawNode(g As Graphics, element As CodeElement)`

- **`DrawNodes`** Private Sub
  - **Parameters:** `g As Graphics, elements() As CodeElement`
  - **Signature:** `Private Sub DrawNodes(g As Graphics, elements() As CodeElement)`

- **`DrawNodeWithOffset`** Private Sub
  - **Parameters:** `g As Graphics, element As CodeElement, pageYOffset As Integer`
  - **Signature:** `Private Sub DrawNodeWithOffset(g As Graphics, element As CodeElement, pageYOffset As Integer)`

- **`GenerateFlowchartPDF`** Public Function
  - **Parameters:** `elements() As CodeElement, outputFile As FolderItem, layoutType As String, includeRelationships As Boolean`
  - **Returns:** `Boolean`
  - **Signature:** `Public Function GenerateFlowchartPDF(elements() As CodeElement, outputFile As FolderItem, layoutType As String, includeRelationships As Boolean) As Boolean`

- **`GetMaxY`** Private Function
  - **Parameters:** `element As CodeElement`
  - **Returns:** `Integer`
  - **Signature:** `Private Function GetMaxY(element As CodeElement) As Integer`

- **`IsElementOnPage`** Private Function
  - **Parameters:** `element As CodeElement, pageYOffset As Integer, pageHeight As Integer`
  - **Returns:** `Boolean`
  - **Signature:** `Private Function IsElementOnPage(element As CodeElement, pageYOffset As Integer, pageHeight As Integer) As Boolean`

- **`LayoutElementHierarchy`** Private Sub
  - **Parameters:** `element As CodeElement, level As Integer, startY As Integer, availableWidth As Integer`
  - **Signature:** `Private Sub LayoutElementHierarchy(element As CodeElement, level As Integer, startY As Integer, availableWidth As Integer)`

#### Events

None

---

### ReportGenerator

#### Properties

None

#### Methods

- **`CalculateAggregateStats`** Private Function
  - **Parameters:** `methodMetrics() As MethodMetrics`
  - **Returns:** `Dictionary`
  - **Signature:** `Private Function CalculateAggregateStats(methodMetrics() As MethodMetrics) As Dictionary`

- **`CalculateCallChainDepth`** Private Function
  - **Parameters:** `element As CodeElement, visited As Dictionary, currentDepth As Integer`
  - **Returns:** `Integer`
  - **Signature:** `Private Function CalculateCallChainDepth(element As CodeElement, visited As Dictionary, currentDepth As Integer) As Integer`

- **`CalculateCyclomaticComplexity`** Private Function
  - **Parameters:** `code As String`
  - **Returns:** `Integer`
  - **Signature:** `Private Function CalculateCyclomaticComplexity(code As String) As Integer`

- **`CalculatePDFHeight`** Private Function
  - **Parameters:** `analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double`
  - **Returns:** `Integer`
  - **Signature:** `Private Function CalculatePDFHeight(analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double) As Integer`

- **`CountOccurrences`** Private Function
  - **Parameters:** `text As String, searchFor As String`
  - **Returns:** `Integer`
  - **Signature:** `Private Function CountOccurrences(text As String, searchFor As String) As Integer`

- **`DrawCenteredText`** Private Sub
  - **Parameters:** `g As Graphics, text As String, centerX As Double, y As Double, maxWidth As Double`
  - **Signature:** `Private Sub DrawCenteredText(g As Graphics, text As String, centerX As Double, y As Double, maxWidth As Double)`

- **`GenerateAnalysisReportPDF`** Public Function
  - **Parameters:** `analyzer As ProjectAnalyzer, saveFile As FolderItem`
  - **Returns:** `Boolean`
  - **Signature:** `Public Function GenerateAnalysisReportPDF(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean`

- **`GetMethodMetrics`** Private Function
  - **Parameters:** `methods() As CodeElement`
  - **Returns:** `MethodMetrics()`
  - **Signature:** `Private Function GetMethodMetrics(methods() As CodeElement) As MethodMetrics()`

- **`RenderComplexityMetrics`** Private Function
  - **Parameters:** `g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderComplexityMetrics(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderFooter`** Private Function
  - **Parameters:** `g As Graphics, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderFooter(g As Graphics, pageWidth As Integer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderHeader`** Private Function
  - **Parameters:** `g As Graphics, pageWidth As Integer, margin As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderHeader(g As Graphics, pageWidth As Integer, margin As Double, yPos As Double) As Double`

- **`RenderRelationships`** Private Function
  - **Parameters:** `g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderRelationships(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderSampleRelationships`** Private Function
  - **Parameters:** `g As Graphics, allMethods() As CodeElement, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderSampleRelationships(g As Graphics, allMethods() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderSummary`** Private Function
  - **Parameters:** `g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderSummary(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderTopComplexMethods`** Private Function
  - **Parameters:** `g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderTopComplexMethods(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderUnusedByType`** Private Function
  - **Parameters:** `g As Graphics, unusedElements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderUnusedByType(g As Graphics, unusedElements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderUnusedElements`** Private Function
  - **Parameters:** `g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderUnusedElements(g As Graphics, analyzer As ProjectAnalyzer, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`RenderUnusedTypeList`** Private Function
  - **Parameters:** `g As Graphics, typeName As String, elements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double`
  - **Returns:** `Double`
  - **Signature:** `Private Function RenderUnusedTypeList(g As Graphics, typeName As String, elements() As CodeElement, margin As Double, lineHeight As Double, yPos As Double) As Double`

- **`SortMethodsByComplexity`** Private Function
  - **Parameters:** `metrics() As MethodMetrics`
  - **Returns:** `MethodMetrics()`
  - **Signature:** `Private Function SortMethodsByComplexity(metrics() As MethodMetrics) As MethodMetrics()`

#### Events

None

---

### MethodMetrics

#### Properties

- **`CallChainDepth`** Public Integer

- **`CallDepth`** Public Integer

- **`Complexity`** Public Integer

- **`Element`** Public CodeElement

- **`LinesOfCode`** Public Integer

- **`MethodName`** Public String

#### Methods

None

#### Events

None

---

### CodeCleanWindowHelpers

**Description:** Module containing shared methods and properties

#### Properties

None

#### Methods

- **`BuildRelationshipSection`** Private Function
  - **Parameters:** `analyzer As ProjectAnalyzer`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildRelationshipSection(analyzer As ProjectAnalyzer) As String`

- **`BuildReportFooter`** Private Function
  - **Returns:** `String`
  - **Signature:** `Private Function BuildReportFooter() As String`

- **`BuildReportHeader`** Private Function
  - **Returns:** `String`
  - **Signature:** `Private Function BuildReportHeader() As String`

- **`BuildSampleRelationships`** Private Function
  - **Parameters:** `methods() As CodeElement`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildSampleRelationships(methods() As CodeElement) As String`

- **`BuildStatisticsSection`** Private Function
  - **Parameters:** `analyzer As ProjectAnalyzer`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildStatisticsSection(analyzer As ProjectAnalyzer) As String`

- **`BuildUnusedByTypeList`** Private Function
  - **Parameters:** `unused() As CodeElement`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildUnusedByTypeList(unused() As CodeElement) As String`

- **`BuildUnusedElementsSection`** Private Function
  - **Parameters:** `analyzer As ProjectAnalyzer`
  - **Returns:** `String`
  - **Signature:** `Private Function BuildUnusedElementsSection(analyzer As ProjectAnalyzer) As String`

- **`CalculateRelationshipStats`** Private Function
  - **Parameters:** `methods() As CodeElement`
  - **Returns:** `Dictionary`
  - **Signature:** `Private Function CalculateRelationshipStats(methods() As CodeElement) As Dictionary`

- **`FormatElementList`** Private Function
  - **Parameters:** `elements() As CodeElement`
  - **Returns:** `String`
  - **Signature:** `Private Function FormatElementList(elements() As CodeElement) As String`

- **`GenerateTextReport`** Public Sub
  - **Parameters:** `Extends win As CodeCleanWindow`
  - **Signature:** `Public Sub GenerateTextReport(Extends win As CodeCleanWindow)`

- **`GroupElementsByType`** Private Function
  - **Parameters:** `elements() As CodeElement`
  - **Returns:** `Dictionary`
  - **Signature:** `Private Function GroupElementsByType(elements() As CodeElement) As Dictionary`

#### Events

None

---

### CodeCleanWindow

**Inherits from:** `DesktopWindow`

**Description:** Desktop Window containing user interface and event handlers

#### Properties

- **`mAnalyzer`** Public ProjectAnalyzer

- **`mLastScannedFolder`** Private FolderItem

#### Methods

None

#### Events

None

---

## Requirements

- **Xojo:** Latest compatible version

## Installation

1. Clone or download this repository
2. Open the `.xojo_project` file in Xojo
3. Build and run the project

## Usage

[Add specific usage instructions for your application]

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

[Specify your license here]

---
*This README was automatically generated from the Xojo project file on 15/11/2025*
*© Philip Cumpston 28th August 2025 *
