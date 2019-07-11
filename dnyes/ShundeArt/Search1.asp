<%
dim count
set rs7=server.createobject("adodb.recordset")
rs7.Source= "select * from "& db_SmallClass_Table &" order by SmallClassID asc"
rs7.open rs7.Source,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0		
        do while not rs7.eof 
        %>
subcat[<%=count%>] = new Array("<%=rs7("SmallClassName")%>","<%=rs7("BigClassid")%>","<%=rs7("SmallClassID")%>");
        <%
        count = count + 1
        rs7.movenext
        loop
        rs7.close
		set rs7=nothing
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassID.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script>
<br> 
<form method="post" name="myform" action="Result.asp" target="newwindow">
  <p>
<select name="action" size="1">
      <option selected value="">全部内容</option>
      <option value="title">按标题</option>
      <option value="content">按内容</option>
	<option value="editor">按作者</option>
	<option value="about">按关键字</option>
    </select>
  </p>
  <p>
<select name="BigClassid" onChange="changelocation(document.myform.BigClassid.options[document.myform.BigClassid.selectedIndex].value)" size="1">
      <option selected value="">全部大类</option>
      <%         
set rs8=server.CreateObject("ADODB.RecordSet")
rs8.Source="select * from "& db_BigClass_Table &" order by BigClassID"
rs8.Open rs8.Source,conn,1,1
        do while not rs8.eof
        %>
      <option value="<%=rs8("BigClassid")%>"><%=rs8("BigclassName")%></option>
      <%
        rs8.movenext
        loop
        rs8.close
		set rs8=nothing
        %>
    </select>
  </p>
    <select name="SmallClassID">                  
        <option selected value="">全部小类</option>
    </select>  
  <p> 
    <input type="text" name="keyword" size=10 value="关键字"onfocus="if (value =='关键字'){value =''}"onblur="if (value ==''){value='关键字'}" maxlength="50">
    <input type="submit" name="Submit" value="搜索">
  </p>
</form>