<HTML>
<HEAD>
<TITLE>计数器</TITLE>
<SCRIPT  LANGUAGE=javascript>
<!--
function f_AddCount(arID){
window.location.href = "add.asp?imgid=" + arID
}
function IMG1_onclick() {
f_AddCount(1);
}

function IMG2_onclick() {
f_AddCount(2);
}

function IMG3_onclick() {
f_AddCount(3);
}

function IMG4_onclick() {
f_AddCount(4);
}

function IMG5_onclick() {
f_AddCount(5);
}

function IMG6_onclick() {
f_AddCount(6);
}

function IMG7_onclick() {
f_AddCount(7);
}

function IMG8_onclick() {
f_AddCount(8);
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<DIV align=center>
<TABLE border=1 cellPadding=1 cellSpacing=1 width="75%">
  <TR>
    <TD colspan=2>
      <P align=center><FONT size=5>计数器</FONT></P>
      </TD></TR> 
    <TR>
    <TD colspan=2>
      <P align=center><a href="display.asp">显示点击率</a>
      </TD></TR> 
  <TR>
    <TD width="50%">
      <P align=right><IMG alt="" id=IMG1 language=javascript onclick="return IMG1_onclick()" src="picture\sohu.gif" ></P></TD>
    <TD><IMG alt=51job id=IMG5 language=javascript onclick="return IMG5_onclick()" src="picture\51job.JPG" ></TD></TR>
  <TR>
    <TD>
      <P align=right><IMG alt="" id=IMG2 language=javascript onclick="return IMG2_onclick()" src="picture\sina.gif" ></P></TD>
    <TD>
      <P align=left><IMG alt="" id=IMG6 language=javascript  onclick="return IMG6_onclick()" src="picture\chinahr.gif" ></P></TD></TR>
  <TR>
    <TD>
      <P align=right><IMG alt="" id=IMG3 language=javascript onclick="return IMG3_onclick()" src="picture\263.gif" ></P></TD>
    <TD>
      <P align=left><IMG alt="" id=IMG7 src="picture\chinaren.gif"  LANGUAGE=javascript onclick="return IMG7_onclick()"></P></TD></TR>
  <TR>
    <TD>
      <P align=right><IMG alt="" id=IMG4 language=javascript onclick="return IMG4_onclick()" src="picture\china.gif" ></P></TD>
    <TD>
      <P align=left><IMG alt="" id=IMG8 src="picture\yahoo.JPG"  LANGUAGE=javascript onclick="return IMG8_onclick()"></P></TD></TR></TABLE></DIV><form name=form method="post" action="add.asp" align="center">
</BODY>
</HTML>