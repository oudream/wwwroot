<HTML><HEAD><TITLE>插入Flash动画</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="site.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
	window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value+"*"+j.value
	window.close();
</SCRIPT>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</HEAD>
<BODY bgColor=menu> 
<form name=flash> 
<br> 
<table border="0" cellspacing="1" cellpadding="0"> 
<tr> 
	<td>&nbsp;</td> 
	<td>
		<fieldset> 
		<legend>FLASH动画定义</legend> 
		<table width="329" border="0" cellpadding="0" cellspacing="1"> 
			<tr> 
				<td width="65" height="25" align="center" valign="middle">动画类型</td> 
				<td height="25" colspan="2"><iframe name="ad" frameborder=0 width=100% height=22 scrolling=no src=upme1.htm></iframe></td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">动画地址</td> 
				<td height="25" colspan="2">
					<input id=a name=a value=''  title='可直接输入网上Flash动画的地址，或由上传程序自动产生Flash动画地址' size="20">
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">动画注释</td> 
				<td height="25" colspan="2">
					<input id=h name=h  title='该注释将自动显示在动画的下方' size="20"> 
					<input id=j type="hidden" name=j value=''  title='可直接输入网上Flash动画的地址，或由上传程序自动产生Flash动画地址' size="20">
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">立即播放</td> 
				<td height="25" colspan="1">
					<select name="select" id=b> 
						<option value="-1">是
						<option value="0">否
					</select>
					</td>
					<td height="25" colspan="1">
					循环播放
					<select id=c>
						<option value="-1">是
						<option value="0">否
					</select>
					<br>
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">显示菜单</td> 
				<td height="25">
					<select name="select" id=d > 
						<option value="-1">是
						<option value="0">否
					</select>
				</td> 
				<td height="25">文件位置
					<select name="select" id=g > 
						<option value="left">居左
						<option value="center">居中
						<option value="right">居右
					</select>
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">动画高度</td> 
				<td width="62" height="25">
					<input name="Input" id=f  ONKEYPRESS="event.returnValue=IsDigit();" value='375' size="10">
				</td> 
				<td width="198" height="25">动画宽度
					<input name="Input" id=e  ONKEYPRESS="event.returnValue=IsDigit();" value='500'  size="10">
				</td> 
			</tr> 
			</table> 
			</fieldset>
		</td> 
		<td>&nbsp;</td> 
		<td align="center" valign="middle">
			<input type=button value=' 插 入  ' id=Ok title='点击“插入”按钮，在编辑器中插入该Flash动画' > 
			<br> 
			<br> 
			<br> 
			<input class=button type=button onClick='window.close();' value=" 取　消 " name="button">
		</td> 
	</tr> 
</table> 
</form> 
</html>