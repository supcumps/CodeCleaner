#tag Class
Protected Class CodeSmell
	#tag Property, Flags = &h0
		SmellType As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Severity As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Element As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		Description As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Details As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Recommendation As String
	#tag EndProperty

	#tag Property, Flags = &h0
		MetricValue As Double
	#tag EndProperty

	#tag Method, Flags = &h0
		Function GetSeverityColor() As Color
		  Select Case Severity
		  Case "CRITICAL"
		    Return Color.RGB(200, 0, 0)  // Red
		  Case "HIGH"
		    Return Color.RGB(220, 100, 0)  // Orange
		  Case "MEDIUM"
		    Return Color.RGB(180, 180, 0)  // Yellow
		  Case "LOW"
		    Return Color.RGB(100, 150, 100)  // Green-gray
		  Else
		    Return Color.Black
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSeverityEmoji() As String
		  Select Case Severity
		  Case "CRITICAL"
		    Return "🔴"
		  Case "HIGH"
		    Return "🟠"
		  Case "MEDIUM"
		    Return "🟡"
		  Case "LOW"
		    Return "🔵"
		  Else
		    Return "⚪"
		  End Select
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
			Name="SmellType"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Severity"
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
			Name="Details"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Recommendation"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="MetricValue"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Double"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
