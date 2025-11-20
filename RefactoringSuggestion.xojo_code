#tag Class
Protected Class RefactoringSuggestion
	#tag Method, Flags = &h0
		Sub Constructor(targetElement As CodeElement, suggestionCategory As String, suggestionPriority As String)
		  // Public Sub Constructor(targetElement As CodeElement, suggestionCategory As String, suggestionPriority As String)
		  Element = targetElement
		  Category = suggestionCategory
		  Priority = suggestionPriority
		  Redim Suggestions(-1)  // Empty array
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Category As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Description As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Element As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		Priority As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Suggestions() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Title As String
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
	#tag EndViewBehavior
End Class
#tag EndClass
