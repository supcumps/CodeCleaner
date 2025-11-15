#tag Class
Protected Class FlowchartGenerator
	#tag Method, Flags = &h0
		Sub CalculateCompactLayout(elements() As CodeElement)
		  // Compact grid layout with reduced spacing and more items per row
		  Var x As Integer = PageMargin
		  Var y As Integer = PageMargin
		  Var maxHeightInRow As Integer = 0
		  Var itemsPerRow As Integer = 5  // More items per row
		  Var itemCount As Integer = 0
		  Var compactHorizontalSpacing As Integer = 30  // Reduced from 50
		  Var compactVerticalSpacing As Integer = 50    // Reduced from 80
		  
		  For Each element As CodeElement In elements
		    element.X = x
		    element.Y = y
		    
		    itemCount = itemCount + 1
		    x = x + element.Width + compactHorizontalSpacing
		    
		    If element.Height > maxHeightInRow Then
		      maxHeightInRow = element.Height
		    End If
		    
		    // Move to next row
		    If itemCount Mod itemsPerRow = 0 Then
		      x = PageMargin
		      y = y + maxHeightInRow + compactVerticalSpacing
		      maxHeightInRow = 0
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CalculateHierarchicalLayout(elements() As CodeElement)
		  // Group elements by type for hierarchical arrangement
		  Var modules() As CodeElement
		  Var classes() As CodeElement
		  Var methods() As CodeElement
		  
		  For Each element As CodeElement In elements
		    Select Case element.ElementType
		    Case "MODULE"
		      modules.Add(element)
		    Case "CLASS"
		      classes.Add(element)
		    Case "METHOD"
		      methods.Add(element)
		    End Select
		  Next
		  
		  Var currentY As Integer = PageMargin
		  
		  // Layout modules at top level
		  For Each moduleElement As CodeElement In modules
		    LayoutElementHierarchy(moduleElement, 0, currentY, PageWidth - (2 * PageMargin))
		    currentY = GetMaxY(moduleElement) + VerticalSpacing
		  Next
		  
		  // Layout standalone classes
		  For Each classElement As CodeElement In classes
		    If classElement.Parent = Nil Then
		      LayoutElementHierarchy(classElement, 0, currentY, PageWidth - (2 * PageMargin))
		      currentY = GetMaxY(classElement) + VerticalSpacing
		    End If
		  Next
		  
		  // Layout standalone methods (unlikely but possible)
		  For Each method As CodeElement In methods
		    If method.Parent = Nil Then
		      method.X = PageMargin
		      method.Y = currentY
		      currentY = currentY + method.Height + VerticalSpacing
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CalculateSimpleLayout(elements() As CodeElement)
		  // Simple grid layout
		  Var x As Integer = PageMargin
		  Var y As Integer = PageMargin
		  Var maxHeightInRow As Integer = 0
		  Var itemsPerRow As Integer = 4
		  Var itemCount As Integer = 0
		  
		  For Each element As CodeElement In elements
		    element.X = x
		    element.Y = y
		    
		    itemCount = itemCount + 1
		    x = x + element.Width + HorizontalSpacing
		    
		    If element.Height > maxHeightInRow Then
		      maxHeightInRow = element.Height
		    End If
		    
		    // Move to next row
		    If itemCount Mod itemsPerRow = 0 Then
		      x = PageMargin
		      y = y + maxHeightInRow + VerticalSpacing
		      maxHeightInRow = 0
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  // Initialize with default values
		  PageWidth = 842  // A4 landscape
		  PageHeight = 595
		  PageMargin = 30
		  HorizontalSpacing = 50
		  VerticalSpacing = 80
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawArrow(g As Graphics, fromElement As CodeElement, toElement As CodeElement)
		  // Calculate connection points (center bottom to center top)
		  Var x1 As Integer = fromElement.X + (fromElement.Width / 2)
		  Var y1 As Integer = fromElement.Y + fromElement.Height
		  Var x2 As Integer = toElement.X + (toElement.Width / 2)
		  Var y2 As Integer = toElement.Y
		  
		  // Draw thicker line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.PenSize = 2
		  g.DrawLine(x1, y1, x2, y2)
		  
		  // Draw larger filled arrowhead triangle
		  Var arrowSize As Integer = 15  // Increased from 10
		  Var angle As Double = ATan2(y2 - y1, x2 - x1)
		  
		  // Calculate three points of the arrowhead triangle
		  Var tipX1 As Double = x2 - arrowSize * Cos(angle - 0.4)
		  Var tipY1 As Double = y2 - arrowSize * Sin(angle - 0.4)
		  Var tipX2 As Double = x2 - arrowSize * Cos(angle + 0.4)
		  Var tipY2 As Double = y2 - arrowSize * Sin(angle + 0.4)
		  
		  // Create path for the triangle
		  Var arrowPath As New GraphicsPath
		  arrowPath.MoveToPoint(x2, y2)
		  arrowPath.AddLineToPoint(tipX1, tipY1)
		  arrowPath.AddLineToPoint(tipX2, tipY2)
		  arrowPath.AddLineToPoint(x2, y2)
		  
		  // Draw filled triangle in darker color
		  g.DrawingColor = Color.RGB(50, 50, 50)
		  g.FillPath(arrowPath)
		  
		  // Draw border around arrowhead for better visibility
		  g.DrawingColor = Color.Black
		  g.PenSize = 1
		  g.DrawPath(arrowPath)
		  
		  // Reset pen size
		  g.PenSize = 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawArrowWithOffset(g As Graphics, fromElement As CodeElement, toElement As CodeElement, pageYOffset As Integer)
		  Var x1 As Integer = fromElement.X + (fromElement.Width / 2)
		  Var y1 As Integer = (fromElement.Y - pageYOffset) + fromElement.Height
		  Var x2 As Integer = toElement.X + (toElement.Width / 2)
		  Var y2 As Integer = toElement.Y - pageYOffset
		  
		  // Draw thicker line
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.PenSize = 2
		  g.DrawLine(x1, y1, x2, y2)
		  
		  // Draw larger filled arrowhead triangle
		  Var arrowSize As Integer = 15  // Increased from 10
		  Var angle As Double = ATan2(y2 - y1, x2 - x1)
		  
		  // Calculate three points of the arrowhead triangle
		  Var tipX1 As Double = x2 - arrowSize * Cos(angle - 0.4)
		  Var tipY1 As Double = y2 - arrowSize * Sin(angle - 0.4)
		  Var tipX2 As Double = x2 - arrowSize * Cos(angle + 0.4)
		  Var tipY2 As Double = y2 - arrowSize * Sin(angle + 0.4)
		  
		  // Create path for the triangle
		  Var arrowPath As New GraphicsPath
		  arrowPath.MoveToPoint(x2, y2)
		  arrowPath.AddLineToPoint(tipX1, tipY1)
		  arrowPath.AddLineToPoint(tipX2, tipY2)
		  arrowPath.AddLineToPoint(x2, y2)
		  
		  // Draw filled triangle in darker color
		  g.DrawingColor = Color.RGB(50, 50, 50)
		  g.FillPath(arrowPath)
		  
		  // Draw border around arrowhead for better visibility
		  g.DrawingColor = Color.Black
		  g.PenSize = 1
		  g.DrawPath(arrowPath)
		  
		  // Reset pen size
		  g.PenSize = 1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawConnections(g As Graphics, elements() As CodeElement)
		  // Draw all relationships
		  For Each element As CodeElement In elements
		    For Each target As CodeElement In element.CallsTo
		      DrawArrow(g, element, target)
		    Next
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawLegend(g As Graphics)
		  Var legendX As Integer = PageWidth - 220
		  Var legendY As Integer = 40
		  Var boxSize As Integer = 12
		  
		  g.FontSize = 8
		  g.Bold = False
		  
		  // Used CLASS
		  g.DrawingColor = Color.RGB(200, 220, 255)
		  g.FillRectangle(legendX, legendY, boxSize, boxSize)
		  g.DrawingColor = Color.Black
		  g.DrawText("Class (used)", legendX + boxSize + 5, legendY + 10)
		  
		  legendY = legendY + 18
		  
		  // Used METHOD
		  g.DrawingColor = Color.RGB(220, 255, 220)
		  g.FillRectangle(legendX, legendY, boxSize, boxSize)
		  g.DrawingColor = Color.Black
		  g.DrawText("Method (used)", legendX + boxSize + 5, legendY + 10)
		  
		  legendY = legendY + 18
		  
		  // Unused
		  g.DrawingColor = Color.RGB(255, 200, 200)
		  g.FillRectangle(legendX, legendY, boxSize, boxSize)
		  g.DrawingColor = Color.Black
		  g.DrawText("Unused", legendX + boxSize + 5, legendY + 10)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawNode(g As Graphics, element As CodeElement)
		  // Color-code by type and usage
		  Select Case element.ElementType
		  Case "CLASS"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(200, 220, 255) // Light blue
		    Else
		      g.DrawingColor = Color.RGB(255, 200, 200) // Light red (unused)
		    End If
		    
		  Case "METHOD"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(220, 255, 220) // Light green
		    Else
		      g.DrawingColor = Color.RGB(255, 230, 200) // Light orange (unused)
		    End If
		    
		  Case "MODULE"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(255, 240, 200) // Light yellow
		    Else
		      g.DrawingColor = Color.RGB(255, 200, 200) // Light red (unused)
		    End If
		    
		  Case "PROPERTY", "VARIABLE"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(240, 220, 255) // Light purple
		    Else
		      g.DrawingColor = Color.RGB(255, 220, 220) // Light pink (unused)
		    End If
		    
		  Else
		    g.DrawingColor = Color.RGB(240, 240, 240) // Light gray
		  End Select
		  
		  // Fill rounded rectangle
		  g.FillRoundRectangle(element.X, element.Y, element.Width, element.Height, 8, 8)
		  
		  // Draw border
		  g.DrawingColor = Color.RGB(80, 80, 80)
		  g.DrawRoundRectangle(element.X, element.Y, element.Width, element.Height, 8, 8)
		  
		  // Draw text - element name
		  g.DrawingColor = Color.Black
		  g.FontSize = 9
		  g.Bold = True
		  
		  Var textX As Integer = element.X + 8
		  Var textY As Integer = element.Y + 20
		  
		  // Show full context if element has parent
		  Var displayName As String = element.Name
		  If element.ParentClass.Trim <> "" Then
		    displayName = element.ParentClass + "." + element.Name
		  ElseIf element.ModuleName.Trim <> "" Then
		    displayName = element.ModuleName + "." + element.Name
		  End If
		  
		  // Truncate long names
		  If displayName.Length > 28 Then
		    displayName = displayName.Left(25) + "..."
		  End If
		  g.DrawText(displayName, textX, textY)
		  
		  // Draw type label
		  g.FontSize = 7
		  g.Bold = False
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawText(element.ElementType, textX, textY + 12)
		  
		  // Draw usage indicator
		  If Not element.IsUsed Then
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.FontSize = 6
		    g.DrawText("UNUSED", textX, textY + 22)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawNodes(g As Graphics, elements() As CodeElement)
		  For Each element As CodeElement In elements
		    DrawNode(g, element)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DrawNodeWithOffset(g As Graphics, element As CodeElement, pageYOffset As Integer)
		  // Adjust Y coordinate for current page
		  Var adjustedY As Integer = element.Y - pageYOffset
		  
		  // Color-code by type and usage
		  Select Case element.ElementType
		  Case "CLASS"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(200, 220, 255)
		    Else
		      g.DrawingColor = Color.RGB(255, 200, 200)
		    End If
		    
		  Case "METHOD"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(220, 255, 220)
		    Else
		      g.DrawingColor = Color.RGB(255, 230, 200)
		    End If
		    
		  Case "MODULE"
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(255, 240, 200)
		    Else
		      g.DrawingColor = Color.RGB(255, 200, 200)
		    End If
		    
		  Else
		    If element.IsUsed Then
		      g.DrawingColor = Color.RGB(240, 220, 255)
		    Else
		      g.DrawingColor = Color.RGB(255, 220, 220)
		    End If
		  End Select
		  
		  g.FillRoundRectangle(element.X, adjustedY, element.Width, element.Height, 8, 8)
		  
		  g.DrawingColor = Color.RGB(80, 80, 80)
		  g.DrawRoundRectangle(element.X, adjustedY, element.Width, element.Height, 8, 8)
		  
		  g.DrawingColor = Color.Black
		  g.FontSize = 9
		  g.Bold = True
		  
		  Var textX As Integer = element.X + 8
		  Var textY As Integer = adjustedY + 20
		  
		  // Show full context if element has parent
		  Var displayName As String = element.Name
		  If element.ParentClass.Trim <> "" Then
		    displayName = element.ParentClass + "." + element.Name
		  ElseIf element.ModuleName.Trim <> "" Then
		    displayName = element.ModuleName + "." + element.Name
		  End If
		  
		  // Truncate long names
		  If displayName.Length > 28 Then
		    displayName = displayName.Left(25) + "..."
		  End If
		  g.DrawText(displayName, textX, textY)
		  
		  g.FontSize = 7
		  g.Bold = False
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  g.DrawText(element.ElementType, textX, textY + 12)
		  
		  If Not element.IsUsed Then
		    g.DrawingColor = Color.RGB(200, 0, 0)
		    g.FontSize = 6
		    g.DrawText("UNUSED", textX, textY + 22)
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GenerateFlowchartPDF(elements() As CodeElement, outputFile As FolderItem, layoutType As String, includeRelationships As Boolean) As Boolean
		  Try
		    // Calculate layout based on type
		    Select Case layoutType
		    Case "Hierarchical"
		      CalculateHierarchicalLayout(elements)
		    Case "Simple"
		      CalculateSimpleLayout(elements)
		    Case "Compact"
		      CalculateCompactLayout(elements)
		    Else
		      CalculateSimpleLayout(elements)
		    End Select
		    
		    // Determine how tall the content is
		    Var maxY As Integer = 0
		    For Each element As CodeElement In elements
		      Var elementMaxY As Integer = element.Y + element.Height
		      If elementMaxY > maxY Then
		        maxY = elementMaxY
		      End If
		    Next
		    
		    // Add margins
		    Var totalHeight As Integer = maxY + (2 * PageMargin)
		    
		    // Create PDF with custom page size
		    Var pdf As New PDFDocument(PageWidth, totalHeight)
		    Var g As Graphics = pdf.Graphics
		    
		    // Set white background
		    g.DrawingColor = Color.White
		    g.FillRectangle(0, 0, PageWidth, totalHeight)
		    
		    // Draw title at top first
		    g.FontSize = 14
		    g.Bold = True
		    g.DrawingColor = Color.Black
		    g.DrawText("Project Code Flowchart", PageMargin, 20)
		    
		    // Draw legend
		    DrawLegend(g)
		    
		    // Draw connections first (so they appear behind boxes)
		    If includeRelationships Then
		      For Each element As CodeElement In elements
		        For Each target As CodeElement In element.CallsTo
		          DrawArrow(g, element, target)
		        Next
		      Next
		    End If
		    
		    // Draw all nodes
		    For Each element As CodeElement In elements
		      DrawNode(g, element)
		    Next
		    
		    // Save PDF
		    pdf.Save(outputFile)
		    Return True
		    
		  Catch e As RuntimeException
		    System.DebugLog("Error generating PDF: " + e.Message)
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetMaxY(element As CodeElement) As Integer
		  Var maxY As Integer = element.Y + element.Height
		  
		  For Each child As CodeElement In element.Contains
		    Var childMaxY As Integer = GetMaxY(child)
		    If childMaxY > maxY Then
		      maxY = childMaxY
		    End If
		  Next
		  
		  Return maxY
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsElementOnPage(element As CodeElement, pageYOffset As Integer, pageHeight As Integer) As Boolean
		  Var elementTop As Integer = element.Y
		  Var elementBottom As Integer = element.Y + element.Height
		  Var pageTop As Integer = pageYOffset
		  Var pageBottom As Integer = pageYOffset + pageHeight
		  
		  // Check if element overlaps with page
		  Return Not (elementBottom < pageTop Or elementTop > pageBottom)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub LayoutElementHierarchy(element As CodeElement, level As Integer, startY As Integer, availableWidth As Integer)
		  element.LayoutLevel = level
		  element.X = PageMargin + (level * 30)  // Indent by level
		  element.Y = startY
		  
		  Var childY As Integer = startY + element.Height + 20
		  
		  For Each child As CodeElement In element.Contains
		    LayoutElementHierarchy(child, level + 1, childY, availableWidth)
		    childY = GetMaxY(child) + 15
		  Next
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		HorizontalSpacing As Integer = 50
	#tag EndProperty

	#tag Property, Flags = &h0
		PageHeight As Integer = 595
	#tag EndProperty

	#tag Property, Flags = &h0
		PageMargin As Integer = 30
	#tag EndProperty

	#tag Property, Flags = &h0
		PageWidth As Integer = 842
	#tag EndProperty

	#tag Property, Flags = &h0
		VerticalSpacing As Integer = 80
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
			Name="PageWidth"
			Visible=false
			Group="Behavior"
			InitialValue="842"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="PageHeight"
			Visible=false
			Group="Behavior"
			InitialValue="595"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="PageMargin"
			Visible=false
			Group="Behavior"
			InitialValue="30"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HorizontalSpacing"
			Visible=false
			Group="Behavior"
			InitialValue="50"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="VerticalSpacing"
			Visible=false
			Group="Behavior"
			InitialValue="80"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
