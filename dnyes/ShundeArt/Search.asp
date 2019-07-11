<br><form method="post" name="myform" action="Result.asp" target="newwindow">
  
<select name="action" size="1">
      <option selected value="">全部内容</option>
      <option value="title">按标题</option>
      <option value="content">按内容</option>
	<option value="editor">按作者</option>
	<option value="about">按关键字</option>
    </select>
 <br><br>
    <input type="text" name="keyword" size=10 value="关键字"onfocus="if (value =='关键字'){value =''}"onblur="if (value ==''){value='关键字'}" maxlength="50">
    <input type="submit" name="Submit" value="搜索"></form>