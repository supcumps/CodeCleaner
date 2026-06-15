#tag Class
Protected Class App
Inherits DesktopApplication
	#tag Method, Flags = &h0
		Sub UnhandledException(error As RuntimeException)
		  //  UnhandledException(error As RuntimeException)
		  Var msg As String = "Error: " + error.Message + EndOfLine
		  msg = msg + "Error Number: " + Str(error.ErrorNumber) + EndOfLine
		  If error.Stack <> Nil Then
		    msg = msg + "Stack:" + EndOfLine
		    For Each frame As String In error.Stack
		      msg = msg + "  " + frame + EndOfLine
		    Next
		  End If
		  
		  Var f As New FolderItem("/tmp/xmcp_debug.log")
		  Var stream As TextOutputStream = TextOutputStream.Open(f)
		  stream.Write(msg)
		  stream.Close
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = Refactored 14/06/26
		Refactoring Session Complete
		All 10 steps done. Clean build. 10 commits on main.
		CommitWhat8f579e9Remove debug MessageBox calls4ea63ccFix Mid() off-by-one i
		n parameter extraction1b8d28bWire up HotSpot.NestingDepth902080bRemove dead code 
		(methods, PNG helpers, stray property)29b9d37PDF page constants in 
		ReportGenerator5d65d72Remove per-element debug Logger.Log477a71aExtract GraphGenerator class 
		(39KB out of ProjectAnalyzer)db9d236Fix duplicate ScanForRelationships + 
		AllElements inconsistency42d3175Extract CodeParser class 
		(37 methods out of ProjectAnalyzer)60b926dUI audit — Nil guards, caption fixes, indentation cleanup
		ProjectAnalyzer size: ~137KB → ~70KB
		
		New classes: CodeParser, GraphGenerator
		
		Bugs fixed: duplicate ScanForRelationships() call, AllElements direct access, off-by-one Mid(),
		 HotSpot.NestingDepth wiring, misplaced AnalyzeRefactoringOpportunities() in parsing pipeline, 
		graph button crash-on-nil
		
	#tag EndNote

	#tag Note, Name = refactoring completed
		CodeCleaner Refactoring — Complete Session Record
		Period: Two sessions across June 14-15, 2026
		
		Final state: 13 commits ahead of origin, clean build, all features working
		Commits (oldest → newest)
		HashDescription8f579e9Remove debug MessageBox calls from RenderRefactoringSuggestions
		4ea63ccFix off-by-one in Mid() for parameter extraction1b8d28bWire up HotSpot.NestingDepth
		902080bRemove dead code (methods, PNG helpers, stray property)
		29b9d37PDF page constants in ReportGenerator
		5d65d72Remove per-element debug Logger.Log
		477a71aExtract GraphGenerator class (39KB out of ProjectAnalyzer)db
		9d236Fix duplicate ScanForRelationships + AllElements inconsistency
		42d3175Extract CodeParser class (37 methods out of ProjectAnalyzer)
		60b926dUI audit — Nil guards, caption fixes, indentation cleanupa
		603e9eFix NilObjectException — restore ProjectAnalyzer Constructor
		5dc0158Fix 0-elements bug — fix getter methods to use correct propertiesc
		1efe47Fix PDF dimensions (Double→Integer) + fix console.Logger.Log
		
		Architecture after refactoring
		CodeCleanWindow (UI coordinator)
		  ├── mAnalyzer: ProjectAnalyzer  (~70KB, was ~137KB)
		  │     └── delegates parsing to CodeParser
		  ├── CodeParser  (37 parsing methods, owns mElements/ElementLookup)
		  ├── GraphGenerator  (11 graph methods, accepts elements() as parameter)
		  ├── ReportGenerator  (PDF output, constants kPage*/kMargin/kLineHeight)
		  ├── HotSpotsGenerator  (hotspot analysis + NestingDepth)
		  └── CodeCleanWindowHelpers  (text report building)
		
	#tag EndNote

	#tag Note, Name = Refactoring Suggestions
		In priority order:
		
		1. Error handling first — it's a medical app
		Audit every database operation, calculation, and file I/O call for Try/Catch coverage. The ones most likely to fail silently are CreateDatabase, OpenDatabase, and any SQLite row reads. A missing or malformed value feeding into a heparin dose calculation is the worst-case scenario.
		
		2. Named constants for clinical coefficients
		The magic numbers in calculateEstimatedBloodVolume, estimateLBW, and calculateProtamineDose should become named constants with their source documented in a comment. Something like:
		// Nadler 1962 male coefficients
		Const kNadlerMaleWeight = 0.3561
		Const kNadlerMaleHeight = 0.03308
		This is both a safety and an audit trail issue. If a reviewer or regulator ever looks at this code, the provenance of those numbers needs to be immediately obvious.
		
		3. Break up the long methods
		resetAll, HandleSegmentPress, and the Opening handler are the practical targets. The approach I'd suggest is extract logically distinct blocks into private methods with descriptive names — the comments you've already written in those methods are natural cut points. This also makes unit testing individual behaviours feasible.
		
		4. Guard clauses in CreateDatabase
		Flatten the four-level nesting with early returns on failure conditions. This is a low-effort, high-readability gain.
		
		5. Leave the XC modules entirely alone
		No value in touching them. The analyser's complaints are artefacts of its ignorance of the bridge architecture.
		What I'd defer indefinitely:
		The God Class and Feature Envy warnings — refactoring jly_iRate or the UIColor bridge class carries real regression risk for essentially zero clinical benefit.
		The pragmatic sequence is: error handling → named constants → long method decomposition → nesting. That order reflects clinical risk descending to code aesthetics.You said: Sounds good.
		
	#tag EndNote


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoQuit"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowHiDPI"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="BugVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Copyright"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Description"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LastWindowIndex"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MajorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MinorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NonReleaseVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RegionCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="StageCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Version"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="_CurrentEventTime"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ProcessID"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
