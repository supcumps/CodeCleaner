#tag Class
Protected Class GraphGenerator
	#tag Method, Flags = &h21
		Private Sub DrawArrow(g As Graphics, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
		  // Public Sub DrawArrow(g As Graphics, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
		  // Draw line
		  g.DrawLine(x1, y1, x2, y2)
		  
		  // Calculate angle for arrowhead
		  Var dx As Double = x2 - x1
		  Var dy As Double = y2 - y1
		  Var angle As Double = ATan2(dy, dx)
		  
		  // Draw arrowhead (larger for visibility)
		  Var arrowSize As Integer = 12
		  Var arrowAngle As Double = 0.4  // Angle of arrowhead
		  
		  Var ax1 As Integer = x2 - arrowSize * Cos(angle - arrowAngle)
		  Var ay1 As Integer = y2 - arrowSize * Sin(angle - arrowAngle)
		  Var ax2 As Integer = x2 - arrowSize * Cos(angle + arrowAngle)
		  Var ay2 As Integer = y2 - arrowSize * Sin(angle + arrowAngle)
		  
		  g.DrawLine(x2, y2, ax1, ay1)
		  g.DrawLine(x2, y2, ax2, ay2)
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Sub DrawLegend(g As Graphics, x As Integer, y As Integer)
		  // Private Sub DrawLegend(g As Graphics, x As Integer, y As Integer)
		  g.FontSize = 12
		  g.Bold = True
		  g.DrawingColor = Color.RGB(44, 62, 80)
		  g.DrawText("Complexity", x, y)
		  
		  g.FontSize = 10
		  g.Bold = False
		  
		  // Green
		  g.DrawingColor = Color.RGB(46, 204, 113)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 20, 20, 20)
		  g.DrawingColor = Color.Black
		  g.DrawText("Low (1-5)", x + 30, y + 35)
		  
		  // Yellow
		  g.DrawingColor = Color.RGB(241, 196, 15)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 50, 20, 20)
		  g.DrawText("Medium (6-10)", x + 30, y + 65)
		  
		  // Red
		  g.DrawingColor = Color.RGB(231, 76, 60)  // FIXED: was ForeColor
		  g.FillRectangle(x, y + 80, 20, 20)
		  g.DrawText("High (11+)", x + 30, y + 95)
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Sub DrawMethodNodes(g As Graphics, nodes() As MethodNode)
		  // Public Sub DrawMethodNodes(g As Graphics, nodes() As MethodNode)
		  For Each node As MethodNode In nodes
		    Var nodeWidth As Integer = 100
		    Var nodeHeight As Integer = 60
		    
		    // Background - use regular rectangle instead of rounded
		    g.DrawingColor = node.NodeColor
		    g.FillRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Border
		    g.DrawingColor = Color.RGB(52, 73, 94)
		    g.PenSize = 2
		    g.DrawRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Method name (shortened if needed)
		    g.FontSize = 9
		    g.Bold = True
		    g.DrawingColor = Color.White
		    
		    Var displayName As String = node.Name
		    If displayName.Length > 15 Then
		      displayName = displayName.Left(12) + "..."
		    End If
		    
		    Var textWidth As Integer = g.TextWidth(displayName)
		    g.DrawText(displayName, node.X - textWidth\2, node.Y - 5)
		    
		    // Metrics (smaller text)
		    g.FontSize = 7
		    g.Bold = False
		    Var metrics As String = "C:" + node.Complexity.ToString + " L:" + node.LOC.ToString
		    Var metricsWidth As Integer = g.TextWidth(metrics)
		    g.DrawText(metrics, node.X - metricsWidth\2, node.Y + 10)
		  Next
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Sub DrawConnections(g As Graphics, deps() As Dependency, nodes() As MethodNode)
		  // Public Sub DrawConnections(g As Graphics, deps() As Dependency, nodes() As MethodNode)
		  If deps.Count = 0 Then Return
		  
		  g.PenSize = 2  // FIXED: was PenWidth
		  g.DrawingColor = Color.RGB(100, 100, 100)
		  
		  For Each dep As Dependency In deps
		    Var fromNode As MethodNode = Nil
		    Var toNode As MethodNode = Nil
		    
		    For Each node As MethodNode In nodes
		      If node.FullPath = dep.FromPath Then fromNode = node
		      If node.FullPath = dep.ToPath Then toNode = node
		    Next
		    
		    If fromNode <> Nil And toNode <> Nil Then
		      DrawArrow(g, fromNode.X, fromNode.Y, toNode.X, toNode.Y)
		    End If
		  Next
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Sub CalculateGraphLayout(nodes() As MethodNode, deps() As Dependency, width As Integer, height As Integer)
		  ///Public Sub CalculateGraphLayout(nodes() As MethodNode, deps() As Dependency, width As Integer, height As Integer)
		  If nodes.Count = 0 Then Return
		  
		  // Group by class/module
		  Var groups As New Dictionary
		  
		  For Each node As MethodNode In nodes
		    Var parts() As String = node.FullPath.Split(".")
		    Var className As String
		    
		    If parts.Count >= 2 Then
		      className = parts(0)
		    Else
		      className = "Global"
		    End If
		    
		    If Not groups.HasKey(className) Then
		      Var groupArray() As MethodNode
		      groups.Value(className) = groupArray
		    End If
		    
		    Var groupArray() As MethodNode = groups.Value(className)
		    groupArray.Add(node)
		    groups.Value(className) = groupArray
		  Next
		  
		  // Layout parameters - MUCH MORE HORIZONTAL SPACE
		  Var startY As Integer = 120
		  Var rowHeight As Integer = 120  // More vertical space between rows
		  Var currentY As Integer = startY
		  Var nodesPerRow As Integer = 8  // More nodes per row (was 6)
		  Var nodeSpacing As Integer = 180  // More horizontal space (was 130)
		  
		  For Each className As Variant In groups.Keys
		    Var groupNodes() As MethodNode = groups.Value(className)
		    
		    // Calculate how many rows needed for this group
		    Var rows As Integer = Ceil(groupNodes.Count / nodesPerRow)
		    
		    For rowIndex As Integer = 0 To rows - 1
		      Var startIdx As Integer = rowIndex * nodesPerRow
		      Var endIdx As Integer = Min(startIdx + nodesPerRow - 1, groupNodes.Count - 1)
		      Var nodesInRow As Integer = endIdx - startIdx + 1
		      
		      // Center the row
		      Var totalRowWidth As Integer = nodesInRow * nodeSpacing
		      Var startX As Integer = (width - totalRowWidth) \ 2
		      
		      // Position each node in the row
		      For i As Integer = startIdx To endIdx
		        Var posInRow As Integer = i - startIdx
		        groupNodes(i).X = startX + (posInRow * nodeSpacing) + nodeSpacing \ 2
		        groupNodes(i).Y = currentY
		      Next
		      
		      currentY = currentY + rowHeight
		    Next
		    
		    // Extra space between groups
		    currentY = currentY + 30
		  Next
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Function BuildInteractiveHTML(nodesJSON As String, edgesJSON As String) As String
		  // Private Function BuildInteractiveHTML(nodesJSON As String, edgesJSON As String) As String
		  Var html As String = "<!DOCTYPE html>" + EndOfLine
		  html = html + "<html>" + EndOfLine
		  html = html + "<head>" + EndOfLine
		  html = html + "  <meta charset=""utf-8"">" + EndOfLine
		  html = html + "  <title>Interactive Dependency Graph</title>" + EndOfLine
		  html = html + "  <script type=""text/javascript"" src=""https://unpkg.com/vis-network@9.1.6/standalone/umd/vis-network.min.js""></script>" + EndOfLine
		  
		  html = html + "  <style>" + EndOfLine
		  html = html + "    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }" + EndOfLine
		  html = html + "    h1 { margin-top: 0; }" + EndOfLine
		  html = html + "    #controls { background: white; padding: 15px; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); position: sticky; top: 0; z-index: 100; }" + EndOfLine
		  html = html + "    #mynetwork { width: 100%; height: 90vh; border: 1px solid #ddd; background: white; border-radius: 8px; min-height: 800px; }" + EndOfLine
		  html = html + "    .control-group { display: inline-block; margin-right: 20px; margin-bottom: 10px; }" + EndOfLine
		  html = html + "    label { font-weight: bold; margin-right: 8px; }" + EndOfLine
		  html = html + "    input, select, button { padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; }" + EndOfLine
		  html = html + "    button { background: #3498db; color: white; cursor: pointer; border: none; }" + EndOfLine
		  html = html + "    button:hover { background: #2980b9; }" + EndOfLine
		  html = html + "    .legend { margin-top: 10px; }" + EndOfLine
		  html = html + "    .legend-item { display: inline-block; margin-right: 20px; }" + EndOfLine
		  html = html + "    .legend-color { display: inline-block; width: 20px; height: 20px; border-radius: 50%; margin-right: 5px; vertical-align: middle; }" + EndOfLine
		  html = html + "    #info { background: #ecf0f1; padding: 10px; margin-top: 10px; border-radius: 4px; display: none; }" + EndOfLine
		  html = html + "    #error { background: #e74c3c; color: white; padding: 10px; margin-bottom: 10px; border-radius: 4px; display: none; }" + EndOfLine
		  html = html + "  </style>" + EndOfLine
		  html = html + "</head>" + EndOfLine
		  html = html + "<body>" + EndOfLine
		  html = html + "  <h1>Interactive Method Dependency Graph</h1>" + EndOfLine
		  html = html + "  <div id=""error""></div>" + EndOfLine
		  html = html + "  <div id=""controls"">" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <label>Search:</label>" + EndOfLine
		  html = html + "      <input type=""text"" id=""searchBox"" placeholder=""Method name..."" />" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <label>Complexity:</label>" + EndOfLine
		  html = html + "      <select id=""complexityFilter"">" + EndOfLine
		  html = html + "        <option value=""all"" selected>All</option>" + EndOfLine
		  html = html + "        <option value=""low"">Low (1-5)</option>" + EndOfLine
		  html = html + "        <option value=""medium"">Medium (6-10)</option>" + EndOfLine
		  html = html + "        <option value=""high"">High (11+)</option>" + EndOfLine
		  html = html + "      </select>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""control-group"">" + EndOfLine
		  html = html + "      <button onclick=""resetView()"">Reset View</button>" + EndOfLine
		  html = html + "      <button onclick=""fitNetwork()"">Fit All</button>" + EndOfLine
		  html = html + "      <button onclick=""zoomIn()"">Zoom In</button>" + EndOfLine
		  html = html + "      <button onclick=""zoomOut()"">Zoom Out</button>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div class=""legend"">" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#2ecc71""></span>Low Complexity (1-5)</span>" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#f1c40f""></span>Medium Complexity (6-10)</span>" + EndOfLine
		  html = html + "      <span class=""legend-item""><span class=""legend-color"" style=""background:#e74c3c""></span>High Complexity (11+)</span>" + EndOfLine
		  html = html + "    </div>" + EndOfLine
		  html = html + "    <div id=""info""></div>" + EndOfLine
		  html = html + "  </div>" + EndOfLine
		  html = html + "  <div id=""mynetwork""></div>" + EndOfLine
		  
		  html = html + "  <script type=""text/javascript"">" + EndOfLine
		  html = html + "    window.onerror = function(msg, url, line) {" + EndOfLine
		  html = html + "      document.getElementById('error').innerHTML = 'Error: ' + msg + ' (line ' + line + ')';" + EndOfLine
		  html = html + "      document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "      return false;" + EndOfLine
		  html = html + "    };" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    if (typeof vis === 'undefined') {" + EndOfLine
		  html = html + "      document.getElementById('error').innerHTML = 'Failed to load vis.js library. Check internet connection.';" + EndOfLine
		  html = html + "      document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "    } else {" + EndOfLine
		  html = html + "      initializeGraph();" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    var allNodes, allEdges, network;" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function initializeGraph() {" + EndOfLine
		  html = html + "      try {" + EndOfLine
		  html = html + "        allNodes = " + nodesJSON + ";" + EndOfLine
		  html = html + "        allEdges = " + edgesJSON + ";" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        console.Logger.Log('Loaded ' + allNodes.length + ' nodes and ' + allEdges.length + ' edges');" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        var nodes = new vis.DataSet(allNodes);" + EndOfLine
		  html = html + "        var edges = new vis.DataSet(allEdges);" + EndOfLine
		  html = html + "        var container = document.getElementById('mynetwork');" + EndOfLine
		  html = html + "        var data = { nodes: nodes, edges: edges };" + EndOfLine
		  html = html + "        var options = {" + EndOfLine
		  html = html + "          physics: {" + EndOfLine
		  html = html + "            enabled: true," + EndOfLine
		  html = html + "            barnesHut: {" + EndOfLine
		  html = html + "              gravitationalConstant: -5000," + EndOfLine
		  html = html + "              centralGravity: 0.05," + EndOfLine
		  html = html + "              springLength: 250," + EndOfLine
		  html = html + "              springConstant: 0.02," + EndOfLine
		  html = html + "              damping: 0.95," + EndOfLine
		  html = html + "              avoidOverlap: 1" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            stabilization: { iterations: 300 }" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          layout: { randomSeed: 42, improvedLayout: true }," + EndOfLine
		  html = html + "          interaction: {" + EndOfLine
		  html = html + "            hover: true," + EndOfLine
		  html = html + "            tooltipDelay: 100," + EndOfLine
		  html = html + "            zoomView: true," + EndOfLine
		  html = html + "            dragView: true," + EndOfLine
		  html = html + "            navigationButtons: true," + EndOfLine
		  html = html + "            keyboard: true" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          nodes: {" + EndOfLine
		  html = html + "            shape: 'box'," + EndOfLine
		  html = html + "            borderWidth: 2," + EndOfLine
		  html = html + "            borderWidthSelected: 3," + EndOfLine
		  html = html + "            margin: 10" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          edges: {" + EndOfLine
		  html = html + "            width: 2," + EndOfLine
		  html = html + "            arrows: { to: { enabled: true, scaleFactor: 0.5 } }," + EndOfLine
		  html = html + "            smooth: { type: 'cubicBezier', forceDirection: 'horizontal' }" + EndOfLine
		  html = html + "          }" + EndOfLine
		  html = html + "        };" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network = new vis.Network(container, data, options);" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network.once('stabilizationIterationsDone', function() {" + EndOfLine
		  html = html + "          network.setOptions({ physics: false });" + EndOfLine
		  html = html + "          setTimeout(function() {" + EndOfLine
		  html = html + "            network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "          }, 100);" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        network.on('click', function(params) {" + EndOfLine
		  html = html + "          if (params.nodes.length > 0) {" + EndOfLine
		  html = html + "            highlightConnections(params.nodes[0]);" + EndOfLine
		  html = html + "            showNodeInfo(params.nodes[0]);" + EndOfLine
		  html = html + "          } else {" + EndOfLine
		  html = html + "            resetHighlight();" + EndOfLine
		  html = html + "            document.getElementById('info').style.display = 'none';" + EndOfLine
		  html = html + "          }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "        setupControls();" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "      } catch(e) {" + EndOfLine
		  html = html + "        document.getElementById('error').innerHTML = 'Error initializing: ' + e.message;" + EndOfLine
		  html = html + "        document.getElementById('error').style.display = 'block';" + EndOfLine
		  html = html + "        console.error(e);" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function highlightConnections(nodeId) {" + EndOfLine
		  html = html + "      var connectedNodes = network.getConnectedNodes(nodeId);" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      allNodes.forEach(function(node) {" + EndOfLine
		  html = html + "        if (node.id === nodeId) {" + EndOfLine
		  html = html + "          // CLICKED NODE - bright blue with thick border" + EndOfLine
		  html = html + "          network.body.data.nodes.update({" + EndOfLine
		  html = html + "            id: node.id," + EndOfLine
		  html = html + "            opacity: 1," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            borderWidth: 5," + EndOfLine
		  html = html + "            color: {" + EndOfLine
		  html = html + "              background: '#3498db'," + EndOfLine
		  html = html + "              border: '#000000'," + EndOfLine
		  html = html + "              highlight: { background: '#3498db', border: '#000000' }" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            font: { color: 'white', size: 14, bold: true }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else if (connectedNodes.indexOf(node.id) !== -1) {" + EndOfLine
		  html = html + "          // CONNECTED NODES - keep original color but with bright border" + EndOfLine
		  html = html + "          var originalNode = allNodes.find(n => n.id === node.id);" + EndOfLine
		  html = html + "          var originalColor = originalNode.color.background || originalNode.color;" + EndOfLine
		  html = html + "          network.body.data.nodes.update({" + EndOfLine
		  html = html + "            id: node.id," + EndOfLine
		  html = html + "            opacity: 1," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            borderWidth: 4," + EndOfLine
		  html = html + "            color: {" + EndOfLine
		  html = html + "              background: originalColor," + EndOfLine
		  html = html + "              border: '#FFD700'," + EndOfLine
		  html = html + "              highlight: { background: originalColor, border: '#FFD700' }" + EndOfLine
		  html = html + "            }," + EndOfLine
		  html = html + "            font: { color: 'white', size: 12, bold: true }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // HIDE unconnected nodes" + EndOfLine
		  html = html + "          network.body.data.nodes.update({id: node.id, hidden: true});" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      // Update edges - make connected edges bright and thick" + EndOfLine
		  html = html + "      allEdges.forEach(function(edge) {" + EndOfLine
		  html = html + "        if (edge.from === nodeId || edge.to === nodeId) {" + EndOfLine
		  html = html + "          // Edges directly from/to clicked node - BRIGHT BLUE & THICK" + EndOfLine
		  html = html + "          network.body.data.edges.update({" + EndOfLine
		  html = html + "            from: edge.from," + EndOfLine
		  html = html + "            to: edge.to," + EndOfLine
		  html = html + "            hidden: false," + EndOfLine
		  html = html + "            width: 4," + EndOfLine
		  html = html + "            color: { color: '#3498db', highlight: '#3498db' }," + EndOfLine
		  html = html + "            arrows: { to: { enabled: true, scaleFactor: 1 } }" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // Hide other edges" + EndOfLine
		  html = html + "          network.body.data.edges.update({from: edge.from, to: edge.to, hidden: true});" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      " + EndOfLine
		  html = html + "      // Auto-fit to show the subgraph" + EndOfLine
		  html = html + "      setTimeout(function() {" + EndOfLine
		  html = html + "        network.fit({" + EndOfLine
		  html = html + "          nodes: [nodeId].concat(connectedNodes)," + EndOfLine
		  html = html + "          padding: 100," + EndOfLine
		  html = html + "          animation: { duration: 400 }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      }, 100);" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function resetHighlight() {" + EndOfLine
		  html = html + "      allNodes.forEach(function(node) {" + EndOfLine
		  html = html + "        var originalColor = node.color.background || node.color;" + EndOfLine
		  html = html + "        network.body.data.nodes.update({" + EndOfLine
		  html = html + "          id: node.id," + EndOfLine
		  html = html + "          opacity: 1," + EndOfLine
		  html = html + "          hidden: false," + EndOfLine
		  html = html + "          borderWidth: 2," + EndOfLine
		  html = html + "          color: {" + EndOfLine
		  html = html + "            background: originalColor," + EndOfLine
		  html = html + "            border: '#2c3e50'," + EndOfLine
		  html = html + "            highlight: { background: originalColor, border: '#000000' }" + EndOfLine
		  html = html + "          }," + EndOfLine
		  html = html + "          font: { color: 'white', size: 10 }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "      allEdges.forEach(function(edge) {" + EndOfLine
		  html = html + "        network.body.data.edges.update({" + EndOfLine
		  html = html + "          from: edge.from," + EndOfLine
		  html = html + "          to: edge.to," + EndOfLine
		  html = html + "          hidden: false," + EndOfLine
		  html = html + "          width: 2," + EndOfLine
		  html = html + "          color: { color: '#95a5a6', highlight: '#3498db' }," + EndOfLine
		  html = html + "          arrows: { to: { enabled: true, scaleFactor: 0.5 } }" + EndOfLine
		  html = html + "        });" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function showNodeInfo(nodeId) {" + EndOfLine
		  html = html + "      var node = allNodes.find(n => n.id === nodeId);" + EndOfLine
		  html = html + "      if (node) {" + EndOfLine
		  html = html + "        var info = document.getElementById('info');" + EndOfLine
		  html = html + "        info.innerHTML = '<strong>' + node.label + '</strong><br>Complexity: ' + node.complexity + ' | LOC: ' + node.loc;" + EndOfLine
		  html = html + "        info.style.display = 'block';" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function setupControls() {" + EndOfLine
		  html = html + "      document.getElementById('searchBox').addEventListener('input', function(e) {" + EndOfLine
		  html = html + "        var term = e.target.value.toLowerCase();" + EndOfLine
		  html = html + "        if (!term) {" + EndOfLine
		  html = html + "          resetHighlight();" + EndOfLine
		  html = html + "          return;" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "        var matches = allNodes.filter(n => n.label.toLowerCase().includes(term));" + EndOfLine
		  html = html + "        if (matches.length > 0) {" + EndOfLine
		  html = html + "          network.selectNodes([matches[0].id]);" + EndOfLine
		  html = html + "          network.focus(matches[0].id, { scale: 1.2, animation: { duration: 500 } });" + EndOfLine
		  html = html + "          highlightConnections(matches[0].id);" + EndOfLine
		  html = html + "          showNodeInfo(matches[0].id);" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "      document.getElementById('complexityFilter').addEventListener('change', function(e) {" + EndOfLine
		  html = html + "        var filter = e.target.value;" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        if (filter === 'all') {" + EndOfLine
		  html = html + "          // Restore everything" + EndOfLine
		  html = html + "          network.body.data.nodes.clear();" + EndOfLine
		  html = html + "          network.body.data.edges.clear();" + EndOfLine
		  html = html + "          network.body.data.nodes.add(allNodes);" + EndOfLine
		  html = html + "          network.body.data.edges.add(allEdges);" + EndOfLine
		  html = html + "        } else {" + EndOfLine
		  html = html + "          // Filter nodes" + EndOfLine
		  html = html + "          var filtered = allNodes.filter(function(node) {" + EndOfLine
		  html = html + "            if (filter === 'low') return node.complexity <= 5;" + EndOfLine
		  html = html + "            if (filter === 'medium') return node.complexity > 5 && node.complexity <= 10;" + EndOfLine
		  html = html + "            if (filter === 'high') return node.complexity > 10;" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          // Get IDs of filtered nodes" + EndOfLine
		  html = html + "          var filteredIds = filtered.map(n => n.id);" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          // Filter edges to only show connections between filtered nodes" + EndOfLine
		  html = html + "          var filteredEdges = allEdges.filter(function(edge) {" + EndOfLine
		  html = html + "            return filteredIds.indexOf(edge.from) !== -1 && filteredIds.indexOf(edge.to) !== -1;" + EndOfLine
		  html = html + "          });" + EndOfLine
		  html = html + "          " + EndOfLine
		  html = html + "          network.body.data.nodes.clear();" + EndOfLine
		  html = html + "          network.body.data.edges.clear();" + EndOfLine
		  html = html + "          network.body.data.nodes.add(filtered);" + EndOfLine
		  html = html + "          network.body.data.edges.add(filteredEdges);" + EndOfLine
		  html = html + "        }" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        setTimeout(function() {" + EndOfLine
		  html = html + "          network.fit({ padding: 100, animation: { duration: 500 } });" + EndOfLine
		  html = html + "        }, 50);" + EndOfLine
		  html = html + "      });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function resetView() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        // Clear and restore both nodes and edges" + EndOfLine
		  html = html + "        network.body.data.nodes.clear();" + EndOfLine
		  html = html + "        network.body.data.edges.clear();" + EndOfLine
		  html = html + "        network.body.data.nodes.add(allNodes);" + EndOfLine
		  html = html + "        network.body.data.edges.add(allEdges);" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Reset all styling" + EndOfLine
		  html = html + "        resetHighlight();" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Clear search and filter" + EndOfLine
		  html = html + "        document.getElementById('searchBox').value = '';" + EndOfLine
		  html = html + "        document.getElementById('complexityFilter').value = 'all';" + EndOfLine
		  html = html + "        document.getElementById('info').style.display = 'none';" + EndOfLine
		  html = html + "        " + EndOfLine
		  html = html + "        // Re-enable physics briefly to reorganize" + EndOfLine
		  html = html + "        network.setOptions({ physics: { enabled: true } });" + EndOfLine
		  html = html + "        setTimeout(function() {" + EndOfLine
		  html = html + "          network.setOptions({ physics: { enabled: false } });" + EndOfLine
		  html = html + "          network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "        }, 500);" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function fitNetwork() {" + EndOfLine
		  html = html + "      if (network) network.fit({ padding: 50, animation: { duration: 500 } });" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function zoomIn() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        var scale = network.getScale();" + EndOfLine
		  html = html + "        network.moveTo({ scale: scale * 1.2, animation: { duration: 300 } });" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "" + EndOfLine
		  html = html + "    function zoomOut() {" + EndOfLine
		  html = html + "      if (network) {" + EndOfLine
		  html = html + "        var scale = network.getScale();" + EndOfLine
		  html = html + "        network.moveTo({ scale: scale * 0.8, animation: { duration: 300 } });" + EndOfLine
		  html = html + "      }" + EndOfLine
		  html = html + "    }" + EndOfLine
		  html = html + "  </script>" + EndOfLine
		  html = html + "</body>" + EndOfLine
		  html = html + "</html>" + EndOfLine
		  
		  Return html
		  
		End Function
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Function GenerateNodesJSON(nodes() As MethodNode) As String
		  // Private Function GenerateNodesJSON(nodes() As MethodNode) As String
		  Var json As String = "["
		  
		  For i As Integer = 0 To nodes.Count - 1
		    Var node As MethodNode = nodes(i)
		    
		    Var safeName As String = node.Name.ReplaceAll("""", "\""")
		    Var safePath As String = node.FullPath.ReplaceAll("""", "\""")
		    Var uniqueId As String = safePath + "_" + i.ToString
		    
		    // Color by complexity
		    Var colorHex As String
		    If node.Complexity <= 5 Then
		      colorHex = "#2ecc71"
		    ElseIf node.Complexity <= 10 Then
		      colorHex = "#f1c40f"
		    Else
		      colorHex = "#e74c3c"
		    End If
		    
		    // SMALLER size based on LOC
		    Var nodeSize As Integer = 10 + Min(node.LOC, 50) \ 5
		    
		    json = json + "{"
		    json = json + """id"": """ + uniqueId + ""","
		    json = json + """label"": """ + safeName + ""","
		    json = json + """title"": ""<b>" + safeName + "</b><br>Complexity: " + node.Complexity.ToString + "<br>LOC: " + node.LOC.ToString + "<br>Path: " + safePath + ""","
		    json = json + """color"": {""background"": """ + colorHex + """, ""border"": ""#2c3e50"", ""highlight"": {""background"": """ + colorHex + """, ""border"": ""#000000""}},"
		    json = json + """size"": " + nodeSize.ToString + ","
		    json = json + """font"": {""color"": ""white"", ""size"": 10, ""face"": ""Arial""},"  // Smaller font
		    json = json + """borderWidth"": 2,"
		    json = json + """complexity"": " + node.Complexity.ToString + ","
		    json = json + """loc"": " + node.LOC.ToString + ","
		    json = json + """fullPath"": """ + safePath + """"
		    json = json + "}"
		    
		    If i < nodes.Count - 1 Then
		      json = json + ","
		    End If
		  Next
		  
		  json = json + "]"
		  Return json
		  
		End Function
	#tag EndMethod
	#tag Method, Flags = &h21
		Private Function GenerateEdgesJSON(elements() As CodeElement, deps() As Dependency) As String
		  // Private Function GenerateEdgesJSON(elements() As CodeElement, deps() As Dependency) As String
		  Var json As String = "["
		  Var edgeCount As Integer = 0
		  
		  // Build a lookup of fullPath to first matching index in methodNodes
		  Var pathToIndex As New Dictionary
		  
		  Var nodeIndex As Integer = 0
		  For Each element As CodeElement In elements
		    If element.ElementType = "METHOD" Then
		      Var key As String = element.FullPath
		      // Store only first occurrence of each path
		      If Not pathToIndex.HasKey(key) Then
		        pathToIndex.Value(key) = nodeIndex
		      End If
		      nodeIndex = nodeIndex + 1
		    End If
		  Next
		  
		  // Generate edges
		  For i As Integer = 0 To deps.Count - 1
		    Var dep As Dependency = deps(i)
		    
		    // Find indices for from and to
		    If pathToIndex.HasKey(dep.FromPath) And pathToIndex.HasKey(dep.ToPath) Then
		      Var fromIdx As Integer = pathToIndex.Value(dep.FromPath)
		      Var toIdx As Integer = pathToIndex.Value(dep.ToPath)
		      
		      Var safeFrom As String = dep.FromPath.ReplaceAll("""", "\""")
		      Var safeTo As String = dep.ToPath.ReplaceAll("""", "\""")
		      
		      If edgeCount > 0 Then json = json + ","
		      
		      json = json + "{"
		      json = json + """from"": """ + safeFrom + "_" + fromIdx.ToString + ""","
		      json = json + """to"": """ + safeTo + "_" + toIdx.ToString + ""","
		      json = json + """arrows"": ""to"","
		      json = json + """color"": {""color"": ""#95a5a6"", ""highlight"": ""#3498db""},"
		      json = json + """smooth"": {""type"": ""cubicBezier""}"
		      json = json + "}"
		      
		      edgeCount = edgeCount + 1
		    End If
		  Next
		  
		  json = json + "]"
		  Return json
		  
		End Function
	#tag EndMethod
	#tag Method, Flags = &h0
		Sub GenerateDependencyGraphPNG(elements() As CodeElement, savePath As FolderItem)
		  // Public Sub GenerateDependencyGraphPNG(elements() As CodeElement, savePath As FolderItem)
		  // Collect methods and their dependencies
		  Var methodNodes() As MethodNode
		  Var dependencies() As Dependency
		  
		  BuildDependencyData(elements, methodNodes, dependencies)
		  
		  // MUCH WIDER canvas for better layout
		  Var canvasWidth As Integer = 1800  // Increased from 2400
		  Var canvasHeight As Integer = 3000  // Taller to accommodate spreading
		  
		  CalculateGraphLayout(methodNodes, dependencies, canvasWidth, canvasHeight)
		  
		  // Find actual bounds needed
		  Var minX As Integer = 999999
		  Var maxX As Integer = 0
		  Var maxY As Integer = 0
		  
		  For Each node As MethodNode In methodNodes
		    minX = Min(minX, node.X - 60)
		    maxX = Max(maxX, node.X + 60)
		    maxY = Max(maxY, node.Y + 60)
		  Next
		  
		  // Add generous margins
		  Var finalWidth As Integer = maxX - minX + 400
		  Var finalHeight As Integer = maxY + 200
		  
		  // Adjust all nodes to proper position with margins
		  For Each node As MethodNode In methodNodes
		    node.X = node.X - minX + 100
		    node.Y = node.Y + 80
		  Next
		  
		  // Create picture
		  Var pic As New Picture(finalWidth, finalHeight)
		  Var g As Graphics = pic.Graphics
		  
		  // White background
		  g.DrawingColor = Color.White
		  g.FillRectangle(0, 0, finalWidth, finalHeight)
		  
		  // Title
		  g.FontSize = 28  // Bigger title
		  g.Bold = True
		  g.DrawingColor = Color.RGB(44, 62, 80)
		  g.DrawText("Method Dependency Graph", 50, 50)
		  
		  // Draw legend
		  DrawLegend(g, finalWidth - 220, 100)
		  
		  // Draw connections first (behind nodes)
		  g.DrawingColor = Color.RGB(150, 150, 150)  // Lighter gray
		  g.PenSize = 1  // Thinner lines
		  
		  For Each dep As Dependency In dependencies
		    Var fromNode As MethodNode = Nil
		    Var toNode As MethodNode = Nil
		    
		    For Each node As MethodNode In methodNodes
		      If node.FullPath = dep.FromPath Then fromNode = node
		      If node.FullPath = dep.ToPath Then toNode = node
		    Next
		    
		    If fromNode <> Nil And toNode <> Nil Then
		      DrawArrow(g, fromNode.X, fromNode.Y, toNode.X, toNode.Y)
		    End If
		  Next
		  
		  // Draw method nodes - LARGER BOXES
		  For Each node As MethodNode In methodNodes
		    Var nodeWidth As Integer = 140  // Wider (was 100)
		    Var nodeHeight As Integer = 70  // Taller (was 60)
		    
		    // Background
		    g.DrawingColor = node.NodeColor
		    g.FillRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Border
		    g.DrawingColor = Color.RGB(52, 73, 94)
		    g.PenSize = 2
		    g.DrawRectangle(node.X - nodeWidth\2, node.Y - nodeHeight\2, nodeWidth, nodeHeight)
		    
		    // Method name - LARGER FONT
		    g.FontSize = 11  // Bigger (was 9)
		    g.Bold = True
		    g.DrawingColor = Color.White
		    
		    Var displayName As String = node.Name
		    If displayName.Length > 18 Then  // Allow longer names
		      displayName = displayName.Left(15) + "..."
		    End If
		    
		    Var textWidth As Integer = g.TextWidth(displayName)
		    g.DrawText(displayName, node.X - textWidth\2, node.Y - 8)
		    
		    // Metrics - LARGER FONT
		    g.FontSize = 9  // Bigger (was 7)
		    g.Bold = False
		    Var metrics As String = "C:" + node.Complexity.ToString + " L:" + node.LOC.ToString
		    Var metricsWidth As Integer = g.TextWidth(metrics)
		    g.DrawText(metrics, node.X - metricsWidth\2, node.Y + 15)
		  Next
		  
		  // Save as PNG
		  Var outputFile As FolderItem = savePath.Child("dependency_graph.png")
		  pic.Save(outputFile, Picture.Formats.PNG)
		  outputFile.Launch()
		  
		  MessageBox("Dependency graph saved as PNG!")
		  
		End Sub
	#tag EndMethod
	#tag Method, Flags = &h0
		Sub GenerateInteractiveDependencyGraph(elements() As CodeElement, savePath As FolderItem)
		  // Public Sub GenerateInteractiveDependencyGraph(elements() As CodeElement, savePath As FolderItem)
		  // Collect data
		  Var methodNodes() As MethodNode
		  Var dependencies() As Dependency
		  
		  BuildDependencyData(elements, methodNodes, dependencies)
		  
		  // Generate JSON data
		  Var nodesJSON As String = GenerateNodesJSON(methodNodes)
		  Var edgesJSON As String = GenerateEdgesJSON(elements, dependencies)
		  
		  // Build HTML with embedded Vis.js
		  Var html As String = BuildInteractiveHTML(nodesJSON, edgesJSON)
		  
		  // Save and open
		  Var outputFile As FolderItem = savePath.Child("dependency_graph.html")
		  Var stream As TextOutputStream
		  
		  Try
		    stream = TextOutputStream.Create(outputFile)
		    stream.Write(html)
		    stream.Close()
		    outputFile.Launch()
		    MessageBox("Interactive dependency graph created!")
		  Catch e As IOException
		    MessageBox("Error saving file: " + e.Message)
		  End Try
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub BuildDependencyData(elements() As CodeElement, ByRef nodes() As MethodNode, ByRef deps() As Dependency)
		  // Private Sub BuildDependencyData(elements() As CodeElement, ByRef nodes() As MethodNode, ByRef deps() As Dependency)
		  // Build nodes from methods
		  For Each element As CodeElement In elements
		    If element.ElementType = "METHOD" Then
		      Var node As New MethodNode
		      node.Name = element.Name
		      node.FullPath = element.FullPath
		      node.Complexity = element.CyclomaticComplexity
		      node.LOC = element.LinesOfCode
		      node.CalledBy = 0
		      
		      // Color by complexity
		      If node.Complexity <= 5 Then
		        node.NodeColor = Color.RGB(46, 204, 113) // Green
		      ElseIf node.Complexity <= 10 Then
		        node.NodeColor = Color.RGB(241, 196, 15) // Yellow
		      Else
		        node.NodeColor = Color.RGB(231, 76, 60) // Red
		      End If
		      
		      nodes.Add(node)
		    End If
		  Next
		  
		  // Build dependencies from code analysis
		  // Look through each element's code for method calls
		  For Each element As CodeElement In elements
		    If element.ElementType = "METHOD" And element.Code.Trim <> "" Then
		      // Scan the code for calls to other methods
		      For Each targetElement As CodeElement In elements
		        If targetElement.ElementType = "METHOD" And targetElement <> element Then
		          // Check if this method is called in the code
		          If element.Code.IndexOf(targetElement.Name) > -1 Then
		            Var dep As New Dependency
		            dep.FromPath = element.FullPath
		            dep.ToPath = targetElement.FullPath
		            deps.Add(dep)
		          End If
		        End If
		      Next
		    End If
		  Next
		  
		End Sub
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
	#tag EndViewBehavior
End Class
#tag EndClass
