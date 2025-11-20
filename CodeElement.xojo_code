#tag Class
Protected Class CodeElement
	#tag Method, Flags = &h0
		Function CalculateLinesOfCode() As Integer
		  // Public Function CalculateLinesOfCode() As Integer
		  If Code.Trim = "" Then
		    LinesOfCode = 0
		    Return 0
		  End If
		  
		  Var lines() As String = Code.Split(EndOfLine)
		  
		  // Count non-empty, non-comment lines
		  Var count As Integer = 0
		  For Each line As String In lines
		    Var trimmed As String = line.Trim
		    If trimmed <> "" And Not trimmed.BeginsWith("//") And Not trimmed.BeginsWith("'") Then
		      count = count + 1
		    End If
		  Next
		  
		  LinesOfCode = count
		  Return count
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(elementType As String, name As String, fullPath As String, Optional fileName As String = "")
		  Self.ElementType = elementType
		  Self.Name = name
		  Self.FullPath = fullPath
		  Self.FileName = fileName
		  Self.IsUsed = False
		  Self.Code = ""
		  Self.ParameterCount = 0
		  Self.OptionalParameterCount = 0
		  Self.Parameters = ""
		  
		  
		  
		  // Parse module and class from full path
		  If fullPath.Contains(".") Then
		    Var parts() As String = fullPath.Split(".")
		    If parts.Count >= 2 Then
		      Self.ModuleName = parts(0)
		      If parts.Count >= 3 Then
		        Self.ParentClass = parts(1)
		      End If
		    End If
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetComplexityScore() As Integer
		  // Calculate overall complexity score for ranking
		  // Higher score = more complex = potentially more resource-intensive
		  
		  Var score As Integer = 0
		  
		  // Line count contributes to complexity
		  score = score + LineCount
		  
		  // Cyclomatic complexity is weighted heavily (x5)
		  score = score + (CyclomaticComplexity * 5)
		  
		  // Number of outgoing calls (x3)
		  score = score + (CallsTo.Count * 3)
		  
		  // Being called by many methods suggests it's important (x2)
		  score = score + (CalledBy.Count * 2)
		  
		  // Parameter complexity (x2)
		  score = score + (ParameterCount * 2)
		  
		  Return score
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		CalledBy() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		CallsTo() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		Code As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Contains() As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		CyclomaticComplexity As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		ElementType As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FileName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FullPath As String
	#tag EndProperty

	#tag Property, Flags = &h0
		HasDatabaseOperations As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		HasFileOperations As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		HasNetworkOperations As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		HasTryCatch As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		HasTypeConversions As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		Height As Integer = 60
	#tag EndProperty

	#tag Property, Flags = &h0
		IsUsed As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		LayoutLevel As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		LineCount As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		LinesOfCode As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		OptionalParameterCount As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		ParameterCount As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		Parameters As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Parent As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		ParentClass As String
	#tag EndProperty

	#tag Property, Flags = &h0
		RefactoringSuggestions() As RefactoringSuggestion
	#tag EndProperty

	#tag Property, Flags = &h0
		RiskyPatterns() As ErrorPattern
	#tag EndProperty

	#tag Property, Flags = &h0
		Width As Integer = 180
	#tag EndProperty

	#tag Property, Flags = &h0
		X As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Y As Integer
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
		#tag ViewProperty
			Name="ElementType"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FullPath"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ParentClass"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ModuleName"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FileName"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsUsed"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="X"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Y"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=false
			Group="Behavior"
			InitialValue="180"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=false
			Group="Behavior"
			InitialValue="60"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LayoutLevel"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LineCount"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CyclomaticComplexity"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Code"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasTryCatch"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasDatabaseOperations"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasFileOperations"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasNetworkOperations"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasTypeConversions"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ParameterCount"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="OptionalParameterCount"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Parameters"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
