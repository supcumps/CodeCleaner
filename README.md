# Xojo Code Analysis Tool

A comprehensive static code analysis tool for Xojo 2025 API 2 projects that detects code smells, analyzes complexity, tracks relationships, and generates detailed multi-page PDF reports.

## Overview

This tool scans Xojo project files (`.xojo_project`, `.xojo_code`, `.xojo_window`, etc.) to provide deep insights into code quality, architecture, and potential issues. It generates professional PDF reports with visualizations, metrics, and actionable recommendations.

## Key Features

### 1. **Code Smell Detection**
Identifies 153+ code quality issues across multiple categories:
- **God Classes**: Classes with too many responsibilities (>40 methods or >1000 LOC)
- **Long Methods**: Methods exceeding 50 lines of code
- **Deep Nesting**: Code with >3 levels of nesting
- **Magic Numbers**: Hard-coded numeric values without explanation
- **Duplicate Code**: Similar code patterns across methods
- **Large Classes**: Classes with excessive size
- **Long Parameter Lists**: Methods with >5 parameters
- **Feature Envy**: Methods accessing other classes' data more than their own
- **Data Clumps**: Same group of parameters appearing together repeatedly
- **Primitive Obsession**: Overuse of primitive types instead of objects

### 2. **Error Handling Analysis**
- **Missing Error Handling**: File operations, network calls, database queries without Try/Catch
- **Bare Catch Blocks**: Catch blocks without proper exception handling
- **Risky Operations**: Type conversions, file operations, network/database calls without protection

### 3. **Complexity Metrics**
- **Cyclomatic Complexity**: Measures code branching and decision paths
- **Lines of Code (LOC)**: Tracks method and class sizes
- **Nesting Depth**: Identifies deeply nested control structures
- **Parameter Complexity**: Analyzes method parameter counts and optional parameters

### 4. **Relationship Analysis**
- **Method Call Detection**: Tracks which methods call which other methods
- **Dependency Mapping**: Shows relationships between classes and modules
- **Call Chain Analysis**: Identifies longest method call chains
- **Unused Element Detection**: Finds unused methods, classes, and properties

### 5. **Multi-Page PDF Reports**
- **Professional formatting** with proper page breaks
- **Color-coded severity indicators** (Critical, High, Medium, Low)
- **Emoji indicators** for visual quick-reference
- **Wrapped text** for long descriptions
- **Section headers** with continuation across pages
- **Quality scores** with component breakdowns

## Architecture

### Core Classes

#### **ProjectAnalyzer**
The main analysis engine that orchestrates code scanning and analysis.

**Key Properties:**
- `AllElements()` - Complete collection of all code elements
- `ClassElements()` - All class definitions
- `MethodElements()` - All method definitions
- `ModuleElements()` - All module definitions
- `DetectedSmells()` - Array of identified code smells
- `ElementLookup` - Dictionary for fast element lookup
- `mMissingErrorHandling()` - Locations missing error handling
- `mBareCatches()` - Bare catch blocks found
- `mRiskyOperations()` - Risky operations without protection

**Key Methods:**
- `AnalyzeProject(projectFolder As FolderItem)` - Main entry point for analysis
- `ScanFileForRelationships(item As FolderItem)` - Scans individual files
- `ScanForRelationships()` - Analyzes call relationships across all methods
- `DetectMethodCalls(code As String, callingMethod As CodeElement)` - Identifies method calls
- `DetectCodeSmells()` - Runs all code smell detection algorithms
- `CalculateQualityScore()` - Computes overall project quality score
- `GetMethodElements()` - Returns all method elements
- `GetAverageParametersPerMethod()` - Calculates parameter complexity
- `GetMethodsExceedingParameterThreshold(threshold As Integer)` - Finds methods with too many parameters

**Analysis Methods:**
- `DetectGodClasses()`
- `DetectLongMethods()`
- `DetectDeepNesting()`
- `DetectMagicNumbers()`
- `DetectLongParameterLists()`
- `AnalyzeMissingErrorHandling()`
- `DetectFileOperations()`
- `DetectDatabaseOperations()`
- `DetectNetworkOperations()`
- `DetectTypeConversions()`

#### **CodeElement**
Represents any code element (class, method, module, property, etc.)

**Key Properties:**
- `Name` - Element name
- `FullPath` - Complete path (e.g., "ClassName.MethodName")
- `FileName` - Source file name
- `ElementType` - Type of element (Class, Method, Module, Property, etc.)
- `Code` - Complete source code for the element
- `LOC` - Lines of code
- `Complexity` - Cyclomatic complexity score
- `NestingDepth` - Maximum nesting level
- `ParameterCount` - Number of parameters (for methods)
- `OptionalParameterCount` - Number of optional parameters
- `Parameters` - Parameter list as string
- `CallsTo()` - Array of methods this element calls
- `CalledBy()` - Array of methods that call this element
- `IsUsed` - Boolean indicating if element is referenced

**Key Methods:**
- `CalculateComplexity()` - Computes cyclomatic complexity
- `CountLinesOfCode()` - Counts effective LOC
- `CalculateNestingDepth()` - Determines maximum nesting
- `ExtractParameters()` - Parses method parameters

#### **CodeSmell**
Represents a detected code quality issue

**Key Properties:**
- `SmellType` - Type of smell (e.g., "God Class", "Long Method")
- `Severity` - Severity level (CRITICAL, HIGH, MEDIUM, LOW)
- `Description` - Human-readable description
- `Details` - Specific metrics (e.g., "Methods: 43, LOC: 995")
- `Recommendation` - Actionable fix suggestion
- `Element` - Reference to the affected CodeElement
- `MethodName` - Name of affected method (if applicable)
- `MetricValue` - Numeric value of the metric that triggered the smell

**Key Methods:**
- `GetSeverityColor()` - Returns color for severity level
- `GetSeverityEmoji()` - Returns emoji indicator (🔴, 🟠, 🟡, 🟢)

#### **ReportGenerator**
Handles PDF report generation with multi-page support

**Key Methods:**
- `GenerateAnalysisReportPDF(analyzer As ProjectAnalyzer, saveFile As FolderItem) As Boolean` - Main PDF generation
- `CheckPageBreak(g As Graphics, yPos As Double, ...) As Double` - Handles page breaks
- `RenderHeader(g As Graphics, ...) As Double` - Renders report header
- `RenderQualityScore(g As Graphics, ...) As Double` - Renders quality metrics
- `RenderCodeSmellsWithPageBreaks(g As Graphics, ...) As Double` - Renders code smells with pagination
- `RenderSingleCodeSmell(g As Graphics, smell As CodeSmell, ...) As Double` - Renders individual code smell
- `RenderErrorHandlingAnalysisWithPageBreaks(g As Graphics, ...) As Double` - Renders error handling section
- `RenderParameterComplexityDetailsWithPageBreaks(g As Graphics, ...) As Double` - Renders parameter analysis
- `RenderUnusedElementsWithPageBreaks(g As Graphics, ...) As Double` - Renders unused elements
- `RenderTopComplexMethodsWithPageBreaks(g As Graphics, ...) As Double` - Renders complexity rankings
- `WrapText(g As Graphics, text As String, maxWidth As Double) As String()` - Text wrapping utility

**PDF Configuration:**
- Standard page size: 612 x 792 pixels (US Letter at 72 DPI)
- Top margin: 80 pixels
- Bottom margin: 100 pixels (reserve for page breaks)
- New page top margin: 120 pixels
- Line height: 14 pixels (default)

## Usage

### Basic Analysis

```xojo
// Create analyzer
Var analyzer As New ProjectAnalyzer

// Select project folder
Var dlg As New SelectFolderDialog
dlg.Title = "Select Xojo Project Folder"
Var projectFolder As FolderItem = dlg.ShowModal()

If projectFolder <> Nil Then
  // Run analysis
  analyzer.AnalyzeProject(projectFolder)
  
  // Generate PDF report
  Var saveDlg As New SaveFileDialog
  saveDlg.SuggestedFileName = "CodeAnalysisReport.pdf"
  saveDlg.Filter = "application/pdf"
  Var saveFile As FolderItem = saveDlg.ShowModal()
  
  If saveFile <> Nil Then
    Var success As Boolean = ReportGenerator.GenerateAnalysisReportPDF(analyzer, saveFile)
    
    If success Then
      MessageBox("Analysis complete! Report saved successfully.")
    Else
      MessageBox("Error generating report.")
    End If
  End If
End If
```

### Accessing Analysis Results Programmatically

```xojo
// Get quality score
Var qualityScore As Double = analyzer.CalculateQualityScore()

// Get code smells
Var smells() As CodeSmell = analyzer.DetectedSmells

// Get unused elements
Var unusedMethods() As CodeElement
For Each element As CodeElement In analyzer.GetMethodElements()
  If Not element.IsUsed Then
    unusedMethods.Add(element)
  End If
Next

// Get methods with high complexity
For Each method As CodeElement In analyzer.GetMethodElements()
  If method.Complexity > 10 Then
    System.DebugLog(method.FullPath + " has complexity: " + method.Complexity.ToString)
  End If
Next

// Get relationship data
For Each method As CodeElement In analyzer.GetMethodElements()
  System.DebugLog(method.Name + " calls " + method.CallsTo.Count.ToString + " methods")
  System.DebugLog(method.Name + " is called by " + method.CalledBy.Count.ToString + " methods")
Next
```

## Analysis Workflow

1. **File Discovery**: Recursively scans project folder for Xojo source files
2. **XML Parsing**: Parses `.xojo_project`, `.xojo_code`, `.xojo_window`, `.xojo_menu` files
3. **Code Extraction**: Extracts code blocks for each element (class, method, property, etc.)
4. **Element Creation**: Creates `CodeElement` objects for each discovered element
5. **Metrics Calculation**: Computes LOC, complexity, nesting depth, parameters
6. **Code Smell Detection**: Runs all smell detection algorithms
7. **Relationship Analysis**: Detects method calls and builds dependency graph
8. **Quality Scoring**: Calculates overall project quality score
9. **Report Generation**: Creates multi-page PDF with all findings

## Code Smell Severity Levels

### 🔴 CRITICAL
- God Classes (>40 methods or >1000 LOC)
- Methods with >8 complexity
- Missing error handling on critical operations

### 🟠 HIGH  
- Long Methods (>50 LOC)
- Deep Nesting (>3 levels)
- Methods with >7 parameters

### 🟡 MEDIUM
- Magic Numbers
- Duplicate Code patterns
- Large Classes (>30 methods)

### 🟢 LOW
- Minor style issues
- Potential optimizations

## Quality Score Components

The overall quality score (0-100) is computed from:

1. **Code Smell Density** (40% weight)
   - Total smells / Total LOC
   - Penalty for CRITICAL and HIGH severity smells

2. **Complexity** (25% weight)
   - Average cyclomatic complexity
   - Percentage of methods with complexity >10

3. **Error Handling** (20% weight)
   - Percentage of risky operations protected
   - Bare catch block count

4. **Code Organization** (15% weight)
   - Average method length
   - Average class size
   - Parameter complexity

## PDF Report Sections

1. **Header**: Project name, analysis date, quality score
2. **Quality Score Breakdown**: Component scores with visual indicators
3. **Code Smells**: Detailed list with type, location, description, recommendation
4. **Error Handling Analysis**: Missing handlers, bare catches, risky operations
5. **Unused Elements**: Methods, classes, properties not referenced
6. **Complexity Metrics**: Top 10 most complex methods
7. **Parameter Complexity**: Methods with high parameter counts
8. **Relationship Analysis**: Method call statistics and dependencies

## Technical Details

### Xojo 2025 API 2 Compatibility

- Uses `Var` instead of `Dim`
- Uses `MessageBox` for simple messages, `MessageDialog` for confirmations
- String comparison: `string.Trim = ""` instead of `string = Nil`
- Arrays: `myArray()` declaration syntax
- Case-insensitive by default

### Performance Considerations

- **Element Lookup Dictionary**: Fast O(1) element access by FullPath
- **Incremental Analysis**: Only analyzes changed files (future enhancement)
- **Memory Management**: Clears temporary data structures after analysis

### Known Limitations

1. **Comment Detection**: Currently removes `//` comments but may need enhancement for `'` and `Rem` styles
2. **Regex Patterns**: Some complex Xojo syntax may not be fully parsed
3. **External References**: Doesn't analyze framework or plugin code
4. **Dynamic Code**: Cannot detect runtime-generated method calls

## Future Enhancements

### Planned Features
- [ ] Interactive HTML reports with clickable dependency graphs
- [ ] Trend analysis across multiple runs
- [ ] Custom rule configuration
- [ ] IDE integration (Xojo plugin)
- [ ] Code fix suggestions with automated refactoring
- [ ] Export to JSON/CSV for CI/CD integration
- [ ] Comparison reports (before/after refactoring)

### Nice-to-Have
- [ ] Table of contents with page numbers
- [ ] Charts and graphs for metrics
- [ ] Historical trend charts
- [ ] Team collaboration features
- [ ] Git integration for change tracking

## Debugging and Logging

The tool includes extensive debug logging via `System.DebugLog()`:

```xojo
System.DebugLog("=== DetectMethodCalls DEBUG ===")
System.DebugLog("Caller: " + callingMethod.Name)
System.DebugLog("Detected " + callingMethod.CallsTo.Count.ToString + " calls")
```

**Tip**: Consider adding a `DebugMode As Boolean` property to toggle logging on/off.

## Troubleshooting

### Common Issues

**Problem**: "Methods with outgoing calls: 0"
- **Solution**: Ensure `ScanForRelationships()` is called after all files are scanned
- Check that method call detection patterns are working

**Problem**: "Code smells section empty in PDF"
- **Solution**: Verify `DetectCodeSmells()` is called before report generation
- Check that smell detection methods are properly implemented

**Problem**: "PDF content cut off at page break"
- **Solution**: Ensure `CheckPageBreak()` is called before rendering each section
- Verify estimated heights are accurate for complex content

**Problem**: "Type mismatch: not an array"
- **Solution**: Check property declarations include `()` for arrays
- Verify properties are Public if accessed from other classes

## Project Structure

```
ProjectAnalyzer (Main Class)
├── Properties
│   ├── AllElements() As CodeElement
│   ├── ClassElements() As CodeElement
│   ├── MethodElements() As CodeElement
│   ├── ModuleElements() As CodeElement
│   ├── DetectedSmells() As CodeSmell
│   ├── mMissingErrorHandling() As String
│   ├── mBareCatches() As String
│   └── mRiskyOperations() As String
└── Methods
    ├── AnalyzeProject()
    ├── ScanForRelationships()
    ├── DetectMethodCalls()
    ├── DetectCodeSmells()
    └── Calculate*() methods

CodeElement (Data Class)
├── Properties (Name, FullPath, Code, LOC, Complexity, etc.)
└── Methods (CalculateComplexity(), CountLinesOfCode(), etc.)

CodeSmell (Data Class)
├── Properties (SmellType, Severity, Description, etc.)
└── Methods (GetSeverityColor(), GetSeverityEmoji())

ReportGenerator (Utility Class)
└── Methods (GenerateAnalysisReportPDF(), Render*(), etc.)
```

## Credits

Developed for Xojo 2025 API 2 projects.

## Version History

### Current Version
- Multi-page PDF report generation
- 153+ code smell detection
- Method call relationship tracking
- Error handling analysis
- Parameter complexity analysis
- Quality score calculation
- Unused element detection

### Recent Fixes
- Fixed PDF header rendering (missing Return statement)
- Fixed multi-page support with g.NextPage
- Fixed code smell details rendering (Element.FullPath fallback)
- Fixed method call detection (uncommented ScanForRelationships call)
- Fixed parameter order in function signatures
- Fixed array declarations for error handling properties

## License

[Your License Here]

## Contact

Philip - Medical Professional & Software Developer
Specializing in Xojo 2025 development and clinical research applications
