#tag Class
Protected Class QualityScore
	#tag Property, Flags = &h0
		OverallScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		ErrorHandlingScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		ComplexityScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		CodeReuseScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		ParameterScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		DocumentationScore As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		Grade As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ErrorHandlingCoverage As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		AverageComplexity As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		UnusedPercentage As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		AverageParameters As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		DocumentationCoverage As Double
	#tag EndProperty

	#tag Method, Flags = &h0
		Function GetScoreColor() As Color
		  // Return color based on score
		  If OverallScore >= 80 Then
		    Return Color.RGB(0, 150, 0)  // Green - Excellent
		  ElseIf OverallScore >= 60 Then
		    Return Color.RGB(100, 150, 0)  // Yellow-Green - Good
		  ElseIf OverallScore >= 40 Then
		    Return Color.RGB(255, 140, 0)  // Orange - Needs Work
		  Else
		    Return Color.RGB(200, 0, 0)  // Red - Critical
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetScoreEmoji() As String
		  // Return emoji based on score
		  If OverallScore >= 80 Then
		    Return "🌟"  // Excellent
		  ElseIf OverallScore >= 60 Then
		    Return "✅"  // Good
		  ElseIf OverallScore >= 40 Then
		    Return "⚠️"  // Needs Work
		  Else
		    Return "🔴"  // Critical
		  End If
		End Function
	#tag EndMethod


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
			Name="OverallScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ErrorHandlingScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ComplexityScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="CodeReuseScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ParameterScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="DocumentationScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Grade"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ErrorHandlingCoverage"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AverageComplexity"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="UnusedPercentage"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AverageParameters"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="DocumentationCoverage"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
