<%@ Language=VBScript %>
<%Option Explicit%>
<%
	Dim items(10,2), maxItems
	maxItems = 9
	items(0,0) = "节点 1"
	items(1,0) = "节点 1"
	items(2,0) = "节点 1"
	items(3,0) = "节点 2"
	items(4,0) = "节点 2"
	items(5,0) = "节点 2"
	items(6,0) = "节点 3"
	items(7,0) = "节点 3"
	items(8,0) = "节点 3"
	items(9,0) = "节点 3"

	items(0,1) = "子节点 1.1"
	items(1,1) = "子节点 1.2"
	items(2,1) = "子节点 1.3"
	items(3,1) = "子节点 2.1"
	items(4,1) = "子节点 2.2"
	items(5,1) = "子节点 2.3"
	items(6,1) = "子节点 3.1"
	items(7,1) = "子节点 3.2"
	items(8,1) = "子节点 3.3"
	items(9,1) = "子节点 3.4"
%>

<%
	Dim index,currentItem, openItem
	openItem = Request.QueryString("open")
%>
<HTML>
<HEAD>
	<TITLE>ASP提纲</TITLE>
</HEAD>
<BODY>
	<h2>ASP提纲</h2>
	<UL>
	
	<%
		index = 0
		Do while index <= maxItems
			if items(index,0) <> currentItem then 
				If items(index,0) <> openItem then _
					Response.Write("<a href='outline.asp?open=" & _
					Server.URLEncode( items(index,0) ) & "'>"  )
					Response.Write( "<LI>" & items(index,0) & "</a>" )
				End If
				
				currentItem = items(index,0)
				If items(index,0) = openItem then 
					Response.Write("<UL>")
					Do While index <= maxItems AND items(index,0) = openItem
						Response.Write( "<LI>" & items(index,1) )
						index = index + 1
					Loop
					Response.Write("</UL>")
				End If
			index = index + 1
		Loop
	%>
	</UL>
</BODY>
</HTML>