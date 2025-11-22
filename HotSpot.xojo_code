#tag Class
Protected Class HotSpot
	#tag Method, Flags = &h0
		Sub Constructor()
		  // Public Sub Constructor()
		  Redim IssueTypes(-1)
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CalledByCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ComplexityScore As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		HotSpotScore As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ImpactDescription As String
	#tag EndProperty

	#tag Property, Flags = &h0
		IssueCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		IssueTypes() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		LinesOfCode As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		MethodPath As String
	#tag EndProperty

	#tag Property, Flags = &h0
		MissingErrorHandling As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		NestingDepth As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ParameterCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		RiskLevel As String
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
			Name="MethodPath"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RiskLevel"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ImpactDescription"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HotSpotScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ComplexityScore"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NestingDepth"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ParameterCount"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LinesOfCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="IssueCount"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MissingErrorHandling"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
