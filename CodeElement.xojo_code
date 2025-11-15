#tag Class
Protected Class CodeElement
	#tag Method, Flags = &h0
		Sub Constructor(elementType As String, name As String, fullPath As String, Optional fileName As String = "")
		  Self.ElementType = elementType
		  Self.Name = name
		  Self.FullPath = fullPath
		  Self.FileName = fileName
		  Self.IsUsed = False
		  Self.Code = ""  // Initialize empty code
		  
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
		ElementType As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FileName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FullPath As String
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
		ModuleName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Parent As CodeElement
	#tag EndProperty

	#tag Property, Flags = &h0
		ParentClass As String
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
			Name="ModuleName"
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
			Name="Code"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
