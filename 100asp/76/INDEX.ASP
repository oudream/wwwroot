<head>
<script LANGUAGE="javascript">
<!--
function research_onsubmit() {
 if(document.research.Options[0].checked==false &&document.research.Options[1].checked==false&&document.research.Options[2].checked==false)
{
alert("错误：没有选择！")
return false
 }
 }
//-->
</script>
</head>
<body>
<div align=center>
<table width="60%"><tr><TD>您认为明年的CBA哪一支队伍能够夺得总冠军？</td></tr><tr></td>
<form method="post" action="add.asp" name="research" LANGUAGE="javascript" onsubmit="return research_onsubmit()" align="center">
 <BR><input type="radio" value="a" name="Options"> 八一队<br>
 <input type="radio" value="b" name="Options"> 上海队<br>
 <input type="radio" value="c" name="Options">说不清楚<br>
 <div align="center"><input type="submit" value=" 提交 " name="cmdok" class="buttonface"> <A href="display.asp">调查结果</a></div> 
</TD></tr></table>
</div></FORM>
</body>