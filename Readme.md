# CodeAnalyser

A comprehensive static code analysis tool for Xojo projects that identifies unused code, analyzes complexity, tracks method relationships, and provides actionable refactoring suggestions.

## Features

### 🔍 Code Analysis
- **Unused Code Detection**: Identifies unused methods, classes, modules, properties, and variables
- **Relationship Tracking**: Maps method calls and dependencies throughout your project
- **Complexity Metrics**: Calculates cyclomatic complexity, lines of code, and nesting depth
- **Error Handling Analysis**: Detects missing try/catch blocks for risky operations
- **Parameter Analysis**: Identifies methods with too many parameters or excessive ByRef usage

### 📊 Refactoring Suggestions
Provides prioritized, actionable recommendations for:
- **Deep Nesting** (4-11+ levels) - with guard clause examples
- **Too Many Parameters** (6+ params) - suggests parameter objects
- **High Complexity** - identifies methods that are hard to test
- **Long Methods** - flags methods that do too much
- **Missing Error Handling** - highlights risky operations without protection

### 📄 Professional Reports
- **Interactive Text Report**: View results directly in the application window
- **PDF Analysis Report**: Comprehensive analysis with metrics and relationships
- **PDF Refactoring Report**: Prioritized suggestions with checkboxes for tracking
- **Flowchart PDFs**: Visual representation of code relationships (multiple layout options)

## Requirements

- **Xojo 2025** (API 2.0)
- **macOS, Windows, or Linux**
- Access to `.xojo_code` project files

## Installation

1. Clone or download this repository
2. Open `CodeAnalyser.xojo_project` in Xojo 2025
3. Run the project (no additional dependencies required)

## How to Use

### 1. Scan a Project

1. Click **"Scan .Xojo_Project File Folder"**
2. Select your Xojo project folder containing `.xojo_code` files
3. Wait for the analysis to complete (progress shown in window)

### 2. View Results

**In-Window Display:**
- Results appear immediately in the text area
- Shows statistics, unused elements, relationships, and suggestions

### 3. Generate Reports

**Export Text Report:**
- Click **"Export Results"** to save as a text file

**Generate PDF Reports:**
- **"Generate Analysis Report PDF"**: Complete analysis with metrics
- **"Generate Refactoring Report"**: Prioritized suggestions with checkboxes

**Create Flowcharts:**
- Click **"Generate Flowchart PDF"**
- Choose layout style from dropdown:
  - **Simple**: Linear layout
  - **Hierarchical**: Tree-based organization
  - **Compact**: Space-efficient arrangement

## What It Analyzes

### Code Elements Tracked
- **Classes**: All class definitions
- **Modules**: Module declarations
- **Methods/Functions**: All callable code blocks
- **Properties**: Class and module properties
- **Variables**: Local and module-level variables
- **Interfaces**: Interface definitions

### Metrics Calculated

#### Complexity Metrics
- **Lines of Code (LOC)**: Non-empty lines per method
- **Cyclomatic Complexity**: Decision points (if, for, while, case, and, or)
- **Nesting Depth**: Maximum indentation levels
- **Call Depth**: Number of methods called directly
- **Call Chain Depth**: Deepest transitive call chain
- **Parameter Count**: Total parameters and optional parameters

#### Code Quality Issues
1. **Deep Nesting**
   - HIGH: ≥ 6 levels
   - MEDIUM: 4-5 levels
   
2. **Too Many Parameters**
   - HIGH: > 5 parameters (especially with ByRef)
   - MEDIUM: 6 parameters
   - CRITICAL: 6+ ByRef parameters

3. **High Complexity**
   - Based on cyclomatic complexity scoring
   - Methods with many decision points

4. **Long Methods**
   - Methods with excessive lines of code
   - Candidates for extraction

5. **Missing Error Handling**
   - Database operations without try/catch
   - File operations without error handling
   - Network operations without protection
   - Type conversions that could fail

### Relationship Analysis
- **Calls To**: Methods called by each method
- **Called By**: Methods that call each method
- **Call Chains**: Transitive dependency paths
- **Unused Detection**: Elements with zero references

## Report Types

### 1. Analysis Report PDF
**Contains:**
- Project summary and statistics
- Unused elements grouped by type
- Complexity metrics and aggregate statistics
- Parameter complexity analysis
- Top 10 most complex methods
- Sample method relationships

**Best for:** Understanding overall code health and identifying optimization targets

### 2. Refactoring Suggestions PDF
**Contains:**
- Total issues count
- HIGH priority issues (fix first!)
- MEDIUM priority issues
- LOW priority issues
- Detailed suggestions with examples
- Checkboxes for tracking progress

**Best for:** Creating an actionable improvement plan

### 3. Flowchart PDF
**Contains:**
- Visual representation of code structure
- Method nodes with metadata
- Relationship arrows
- Color-coded by element type
- Interactive legend

**Best for:** Understanding code architecture and dependencies

## Technical Details

### Architecture

#### Core Components
1. **ProjectAnalyzer**: Main analysis engine
   - Scans and parses `.xojo_code` files
   - Builds element dictionary and relationship graph
   - Calculates all metrics

2. **ReportGenerator**: PDF creation engine
   - Uses Xojo's PDFDocument API
   - Renders formatted reports with graphics
   - Handles text wrapping and pagination

3. **CodeElement**: Data structure
   - Stores element metadata
   - Tracks relationships and metrics
   - Holds refactoring suggestions

4. **RefactoringSuggestion**: Suggestion data
   - Priority level (HIGH/MEDIUM/LOW)
   - Category and description
   - Actionable recommendations

5. **FlowchartGenerator**: Visualization engine
   - Layout algorithms (simple, hierarchical, compact)
   - Node positioning and rendering
   - Relationship arrow drawing

### Analysis Process

```
1. Scan Project Folder
   ├── Find all .xojo_code files
   └── Parse file structure

2. Extract Code Elements
   ├── Identify classes, modules, methods
   ├── Parse method signatures and parameters
   └── Extract code blocks

3. Build Relationships
   ├── Detect method calls
   ├── Track references
   └── Build dependency graph

4. Analyze Metrics
   ├── Calculate complexity
   ├── Measure nesting depth
   ├── Count parameters
   └── Assess code quality

5. Generate Suggestions
   ├── Deep nesting analysis
   ├── Parameter complexity
   ├── Error handling coverage
   └── Method length assessment

6. Deduplicate & Prioritize
   ├── Remove duplicate suggestions
   ├── Assign priority levels
   └── Sort by importance

7. Generate Reports
   ├── Format for display/export
   ├── Render PDFs
   └── Create visualizations
```

### Key Algorithms

#### Cyclomatic Complexity Calculation
```
Complexity = 1 (base)
  + count(IF, ELSEIF)
  + count(FOR, WHILE, DO)
  + count(CASE statements)
  + count(AND, OR operators)
```

#### Nesting Depth Detection
- Tracks indentation levels in code blocks
- Identifies nested IF/FOR/WHILE/TRY structures
- Reports maximum depth reached

#### Deduplication Strategy
- Creates unique key: `methodName|category|description`
- Uses Dictionary to track seen suggestions
- Preserves first occurrence, discards duplicates
- Applies to both PDF and window display

## Understanding the Reports

### Priority Levels

**🔴 HIGH PRIORITY** - Fix These First!
- Deep nesting (6+ levels)
- Critical parameter issues (6+ ByRef)
- Severe complexity problems
- Missing error handling in critical operations

**🟡 MEDIUM PRIORITY** - Important Improvements
- Moderate nesting (4-5 levels)
- Too many parameters (6)
- Moderate complexity issues
- Code organization opportunities

**🔵 LOW PRIORITY** - Nice to Have
- Minor improvements
- Style consistency
- Optional optimizations

### Reading Suggestions

Each suggestion includes:
1. **Method Path**: Full path to the problematic method
2. **Issue Type**: Category of the problem
3. **Description**: Clear explanation of the issue
4. **Actionable Steps**: Specific recommendations with examples
5. **Example Code**: Before/after refactoring patterns

## Best Practices

### When to Use CodeAnalyser

✅ **Before Major Refactoring**
- Identify technical debt
- Plan improvement roadmap
- Prioritize changes

✅ **During Code Reviews**
- Objective quality metrics
- Standardize feedback
- Track improvements

✅ **Project Health Checks**
- Quarterly code audits
- Team training opportunities
- Architecture documentation

✅ **Before Releases**
- Verify no unused code
- Check complexity hasn't grown
- Ensure error handling coverage

### Interpreting Results

**Unused Code:**
- May be legitimately unused (dead code)
- Could be called dynamically (introspection)
- Might be event handlers (not detected as "called")
- Always verify before deleting

**High Complexity:**
- Not always bad (some domains are complex)
- Consider business logic complexity
- Focus on testability and maintainability

**Refactoring Suggestions:**
- Start with HIGH priority
- Fix one category at a time
- Test after each change
- Use version control

## Limitations

- **Dynamic Calls**: Does not detect runtime method invocation
- **Event Handlers**: May flag event handlers as unused
- **Compiled Methods**: Cannot analyze precompiled libraries
- **External References**: Only analyzes files in scanned folder
- **Xojo-Specific**: Designed for Xojo code syntax only

## Development

### Built With
- **Language**: Xojo (API 2.0)
- **Platform**: Cross-platform (macOS, Windows, Linux)
- **PDF Generation**: Native Xojo PDFDocument
- **Graphics**: Xojo Graphics API

### Project Structure
```
CodeAnalyser/
├── ProjectAnalyzer.xojo_code      # Core analysis engine
├── ReportGenerator.xojo_code      # PDF generation
├── FlowchartGenerator.xojo_code   # Visualization
├── CodeElement.xojo_code          # Data structure
├── RefactoringSuggestion.xojo_code # Suggestion model
├── MethodMetrics.xojo_code        # Metrics container
├── CodeCleanWindow.xojo_window    # Main UI
└── CodeCleanWindowHelpers.xojo_code # Text report generation
```

## Tips & Tricks

### Optimizing Analysis Speed
- Scan only the folders you need
- Exclude build/temporary folders
- Run analysis on clean projects

### Getting Better Suggestions
- Keep methods focused (single responsibility)
- Use meaningful names (aids detection)
- Add error handling proactively
- Limit parameter counts from the start

### Using Reports Effectively
- Print refactoring PDFs for team reviews
- Track checkbox progress over sprints
- Compare reports over time
- Share flowcharts for documentation

## Troubleshooting

**Issue: No suggestions generated**
- Verify .xojo_code files were found
- Check that methods have code (not just declarations)
- Ensure file permissions allow reading

**Issue: PDF appears blank**
- Check Messages panel for errors
- Verify write permissions for output folder
- Try a smaller project first

**Issue: Too many false positives**
- Focus on HIGH priority only initially
- Verify unused code isn't dynamically called
- Adjust thresholds if needed (requires code modification)

## Future Enhancements

Potential improvements:
- Custom threshold configuration
- Historical trend analysis
- Team collaboration features
- Integration with CI/CD pipelines
- Support for other languages
- Automated refactoring assistance

## License

MIT License

Copyright (c) 2025 Philip Cumpston

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Test thoroughly
4. Submit a pull request

## Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Contact: [Your contact information]

## Acknowledgments

Built to help Xojo developers write better, cleaner code through automated analysis and actionable insights.

---

**Made with ❤️ in Xojo**
