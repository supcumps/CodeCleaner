#tag Class
Protected Class ManifestEntry
	#tag Method, Flags = &h0
		Sub Constructor(name As String, fileName As String, elementType As String, sourceFile As FolderItem)
		  Self.Name = name
		  Self.FileName = fileName
		  Self.ElementType = elementType
		  Self.SourceFile = sourceFile
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsResolved() As Boolean
		  // Returns True if the manifest's referenced file actually exists on disk
		  Return Self.SourceFile <> Nil And Self.SourceFile.Exists
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		ElementType As String
	#tag EndProperty

	#tag Property, Flags = &h0
		FileName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceFile As FolderItem
	#tag EndProperty

End Class
#tag EndClass
