<HTML><HEAD>
<TITLE>插入图片</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<link rel="stylesheet" type="text/css" href="site.css">
<SCRIPT language=JavaScript event=onclick for=Ok>
	window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value
	window.close();
</SCRIPT>
</HEAD>
<BODY bgColor=menu> 
<table width="200" border="0" cellspacing="1" cellpadding="0"> 
<FORM name=pic >
<tr> 
	<td>&nbsp;</td> 
	<td>
		<fieldset> 
		<legend>图片定义</legend> 
		<table width="326" border="0" cellspacing="1" cellpadding="0"> 
		<tr> 
			<td width="59" height="25" align="center" valign="middle">图片类型</td> 
			<td width="238" height="25"><IFRAME name=ad src="upme2.htm" frameBorder=0 width=100% scrolling=no height=22></IFRAME> </td> 
		</tr> 
		<tr> 
			<td height="25" align="center" valign="middle">图片地址</td> 
			<td height="25"><INPUT id=a title="可直接输入网上图片的地址，或由上传程序自动产生图片地址" style="WIDTH: 200px" value="" name=a ></td> 
		</tr> 
		<tr> 
			<td height="25" align="center" valign="middle">图片注释</td> 
			<td height="25"><INPUT id=d title=该注释将自动显示在图片的下方  name=d  size="20"> </td> 
		</tr> 
		<tr> 
			<td height="25" align="center" valign="middle">提示文字</td> 
			<td height="25"><INPUT id=b title=当鼠标移到图片上时，将会出现该提示文字   size="20"> </td> 
		</tr> 
		<tr> 
			<td height="25" align="center" valign="middle">特殊效果</td> 
			<td height="25"><INPUT id=f type="hidden" title=可直接输入网上图片的地址，或由上传程序自动产生图片地址  value="" name=f size="20"> 
				<SELECT id=c> 
					<OPTION selected>不应用
					<OPTION value=filter:blur(add=1,direction=14,strength=15)>模糊效果
					<OPTION value=filter:gray>黑白照片
					<OPTION value=filter:flipv>颠倒显示
					<OPTION value=filter:fliph>左右反转
					<OPTION value=filter:xray>X 光照片
					<OPTION value=filter:invert>颜色反转</OPTION> 
				</SELECT> 
&nbsp;&nbsp; 
				<SELECT id=e> 
					<OPTION value="" selected>默认
					<OPTION value=left>居左
					<OPTION value=center>居中
					<OPTION value=right>居右</OPTION> 
				</SELECT>
			</td> 
		</tr> 
		</table> 
		</fieldset>
	</td> 
	<td>&nbsp;</td> 
	<td>
		<INPUT id=Ok title=点击“插入”按钮，在编辑器中插入该图片 type=button value=" 插 入 " >
		<br>
		<br>
		<br>
		<input class=button type=button onClick='window.close();' value=" 取 消 " name="button">
	</td> 
</tr> 
</FORM> 
</table> 
</BODY>
</HTML>