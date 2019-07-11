<!--#include file="css.asp"-->
<html>
<head>
<title>JavaScript对象与数组参考大全</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="top.asp" -->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/2x2.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td><img src="images/1-x.gif" width="754" height="56"></td>
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top"><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD align="center" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid; BORDER-RIGHT: #000000 1px solid"> 
                    <table width="100%" border="0" cellpadding="5" cellspacing="0">
                      <tr>
                        <td><div align="center">JavaScript对象与数组参考大全 </div>
                          <p></p>
                          <p>--------------------------------------------------------------------------------<br>
                            如果你有这个方面的问题请到JAVA&amp;JSP去提问或发表你的意见 </p>
                          <p><br>
                            来源：论坛转载无法确定出处，如有版权问题请与我们联系</p>
                          <p>　　本文列举了各种JavaScript对象与数组,同时包括对上述每一对象或数组所完成工作的简短描述,以及与其相关的属性方法,以及事件处理程序,还注明了该对象或数组的父对象用户同样可能需要参考Online 
                            Companion中的超级文本Object Hierarchy页面(http://www.netscapepress.com/support/javascript/10-9.htm),以便了解这些对象之间是如何相互关联的。</p>
                          <p>　　顺便提一下,记住,这里把所有作为另一对象的子对象的对象看作该对象的属性请参见第十章中与此相关的注解。</p>
                          <p>　　B.1 anchor对象</p>
                          <p>　　使用&lt;A NAME=&gt;标记创建的HTML描点能被一个链接作为目标如果锚点包括HREF=特性,则它也是一个链接对象。</p>
                          <p>　　anchor对象是document对象的一个属性,它本身没有属性方法或者事件处理程序。</p>
                          <p>　　B.2 anchors数组</p>
                          <p>　　anchors数组是document对象的一个属性,是文档内所有anchor对象的一个列表如果anchor也是一个link(链接),则它会同时出现在anchors和links数组中。</p>
                          <p>　　属性</p>
                          <p>　　length 文档内的锚点个数</p>
                          <p>　　B.3 array对象</p>
                          <p>　　array对象是Netscape Navlgator 3.0 beta 3中引入的一个新的对象,因而,它不能在Netscape 
                            2.0中使用它是一个内置对象,而不是其它对象的属性。</p>
                          <p>　　属性</p>
                          <p>　　length 数组中的值个数</p>
                          <p>　　B.4 button对象</p>
                          <p>　　它是form对象的一个属性,使用&lt;INPUT TYPE=&quot;BUTTON&quot;&gt;标记来创建。</p>
                          <p>　　属性</p>
                          <p>　　name HTML标记中的NAME=特性</p>
                          <p>　　value HTML标记中的VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　click 模拟鼠标单击一按钮</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onclick</p>
                          <p>　　B.5 checkbox 对象</p>
                          <p>　　它是form对象的一个属性,使用&lt;INPUT TYPE=&quot;CHECKBOX&quot;&gt;标记来创建。</p>
                          <p>　　属性</p>
                          <p>　　checked 复选框的选择状态</p>
                          <p>　　defaultChecked 标记的CHECKED=特性</p>
                          <p>　　name 标记的NAME=特性</p>
                          <p>　　value 标记的VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　click 模拟鼠标单击按钮</p>
                          <p>　　事件处理程序</p>
                          <p>　　onclick</p>
                          <p>　　B.6 Date对象</p>
                          <p>　　它是一个内置对象——而不是其它对象的属性,允许用户执行各种使用日期和时间的过程。</p>
                          <p>　　方法</p>
                          <p>　　getDate() 查看Date对象并返回日期</p>
                          <p>　　getDay() 返回星期几</p>
                          <p>　　getHours() 返回小时数</p>
                          <p>　　getMinutes() 返回分钟数</p>
                          <p>　　getMonth() 返回月份值</p>
                          <p>　　getSeconds() 返回秒数</p>
                          <p>　　getTime() 返回完整的时间</p>
                          <p>　　getTimezoneoffset() 返回时区偏差值(格林威治平均时间与运行脚本的计算机所处时区设置之间相差的小时数)</p>
                          <p>　　getYear() 返回年份</p>
                          <p>　　parse() 返回在Date字符串中自从1970年1月1日00:00:00以来的毫秒数(Date对象按照毫秒数的形式存储从那时起的日期和时间)但是注意,该方法当前不能正确运行</p>
                          <p>　　setDate() 改变Date对象的日期</p>
                          <p>　　setHours() 改变小时数</p>
                          <p>　　setMinutes() 改变分钟数</p>
                          <p>　　setMonth() 改变月份</p>
                          <p>　　setSeconds() 改变秒数</p>
                          <p>　　setTime() 改变完整的时间</p>
                          <p>　　setYear() 改变年份</p>
                          <p>　　toGMTString() 把Date对象的日期(一个数值)转变成一个GMT时间字符串,返回类似下面的值:Weds,15 
                            June l997 14:02:02 GMT(精确的格式依赖于计算机上所运行的操作系统而变)</p>
                          <p>　　toLocaleString() 把Date对象的日期(一个数值)转变成一个字符串,使用所在计算机上配置使用的特定日期格式</p>
                          <p>　　UTC() 使用Date UTC(年、月、日、时、分、秒),以自从1970年1月1日00:00:00(其中时、分、秒是可选的)以来的毫秒数的形式返回日期</p>
                          <p>　　B.7 document对象</p>
                          <p>　　该对象是window和frames对象的一个属性,是显示于窗口或框架内的一个文档。</p>
                          <p>　　属性</p>
                          <p>　　alinkColor 活动链接的颜色(ALINK)</p>
                          <p>　　anchor 一个HTMI锚点,使用&lt;A NAME=&gt;标记创建(该属性本身也是一个对象)</p>
                          <p>　　anchors array 列出文档锚点对象的数组(&lt;A NAME=&gt;)(该属性本身也是一个对象)</p>
                          <p>　　bgColor 文档的背景颜色(BGCOLOR)</p>
                          <p>　　cookie 存储于cookie.txt文件内的一段信息,它是该文档对象的一个属性</p>
                          <p>　　fgColor 文档的文本颜色(&lt;BODY&gt;标记里的TEXT特性)</p>
                          <p>　　form 文档中的一个窗体(&lt;FORM&gt;)(该属性本身也是一个对象)</p>
                          <p>　　forms anay 按照其出现在文档中的顺序列出窗体对象的一个数组(该属性本身也是一个对象)</p>
                          <p>　　lastModified 文档最后的修改日期</p>
                          <p>　　linkColor 文档的链接的颜色,即&lt;BODY&gt;标记中的LINK特性(链接到用户没有观察到的文档)</p>
                          <p>　　link 文档中的一个&lt;A HREF=&gt;标记(该属性本身也是一个对象)</p>
                          <p>　　links array 文档中link对象的一个数组,按照它们出现在文档中的顺序排列(该属性本身也是一个对象)</p>
                          <p>　　location 当前显示文档的URL。用户不能改变document.location(因为这是当前显示文档的位置)。但是,可以改变window.location 
                            (用其它文档取代当前文档)window.location本身也是一个对象,而document.location不是对象</p>
                          <p>　　referrer 包含链接的文档的URL,用户单击该链接可到达当前文档</p>
                          <p>　　title 文档的标题((TITLE&gt;)</p>
                          <p>　　vlinkColor 指向用户已观察过的文档的链接文本颜色,即&lt;BODY&gt;标记的VLINK特性</p>
                          <p>　　方法</p>
                          <p>　　clear 清除指定文档的内容</p>
                          <p>　　close 关闭文档流</p>
                          <p>　　open 打开文档流</p>
                          <p>　　write 把文本写入文档</p>
                          <p>　　writeln 把文本写入文档,并以换行符结尾</p>
                          <p>　　B.8 elements数组</p>
                          <p>　　它是form对象的一个属性,列举了窗体内各元素的一个数组。</p>
                          <p>　　属性</p>
                          <p>　　1ength 窗体内的元素个数</p>
                          <p>　　B.9 form对象</p>
                          <p>　　它是document对象的一个属性,文档内的一个窗体。</p>
                          <p>　　属性</p>
                          <p>　　action 包含了为一个窗体提交的目标URL的字符串</p>
                          <p>　　button 窗体内的一个按钮,使用&lt;INPUT TYPE=”BUTTON”&gt;标记来创建(该属性本身也是一个对象)</p>
                          <p>　　checkbox 复选框,使用&lt;INPUT TYPE=”CHECKBOX”&gt;标记来创建 
                            (该属性本身也是一个对象)</p>
                          <p>　　elements array 一个数组,按照其出现于窗体内的顺序列举各窗体元素(该属性本身也是一个对象)</p>
                          <p>　　encoding 窗体的MIME编码</p>
                          <p>　　hidden 窗体里的一个隐藏元素(&lt;INPUT TYPE=”HIDDEN”&gt;)。窗体对象的一个属性(该属性本身也是一个对象)</p>
                          <p>　　length 窗体里的元素的个数</p>
                          <p>　　method 输入窗体的数据传送到服务器上的方式,即(FORM)标记中的METHOD特性</p>
                          <p>　　radio 设置在窗体里的单选按钮(&lt;INPUT TYPE=”RADIO”&gt;)(该属性本身也是一个对象)</p>
                          <p>　　reset 窗体里的复位按钮((1NPUT TYPE=”RESET”&gt;)(该属性自身也是一个对象)</p>
                          <p>　　select 窗体里的选择框(&lt;SELECT&gt;)(该属性本身也是一个对象)</p>
                          <p>　　submit 窗体里的提交按钮(&lt;INPUT TYPE=”SUBMIT”&gt;)(该属性本身也是一个对象)</p>
                          <p>　　target 提交窗体后,显示回应信息的窗口的名字</p>
                          <p>　　text 窗体里的文本元素(&lt;INPUT TYPE=”TEXT”&gt;)(该属性本身也是一个对象)</p>
                          <p>　　textarta 窗体里的文本区元素(&lt;TEXTAREA&gt;)(该属性本身也是一个对象)</p>
                          <p>　　方法</p>
                          <p>　　submit 提交窗体(与使用Submit按钮的作用相同)事件处理程序</p>
                          <p>　　onsubmit</p>
                          <p>　　B.10 forms数组</p>
                          <p>　　该数组是document对象的一个属性,即列举了文档内的各窗体的一个数组。</p>
                          <p>　　属性</p>
                          <p>　　length 文档内窗体的个数</p>
                          <p>　　B.11 frame对象</p>
                          <p>　　它是window对象的一个属性,窗口内的一个框架。除了个别例外,frame对象与window对象的作用相同。</p>
                          <p>　　属性</p>
                          <p>　　frames array 列举该框架内的各个子框架的一个数组(该属性本身也是—个对象)</p>
                          <p>　　length 该框架内的框架数</p>
                          <p>　　name 框架的名字(&lt;FRAME&gt;标记里的NAME特性)</p>
                          <p>　　parent 包含本框架的父窗口的同义词</p>
                          <p>　　self 当前框架的同义词</p>
                          <p>　　window 当前框架的同义词</p>
                          <p>　　方法</p>
                          <p>　　clearTimeout() 用来终止setTimeout方法的工作</p>
                          <p>　　setTimeout() 等待指定的毫秒数,然后运行指令</p>
                          <p>　　B.12 frames数组</p>
                          <p>　　它既是window对象,也是frame对象的属性,列举了window或者frame对象内的各框架。</p>
                          <p>　　属性</p>
                          <p>　　length 窗口或框架对象内的框架数</p>
                          <p>　　B.13 hidden对象</p>
                          <p>　　糊为form对象的一个属性,窗体内的一个隐藏元素(&lt;INPUT TYPE=”HIDDEN”&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　name 标记内的名字(NAME特性)</p>
                          <p>　　value 标记内的VALUE=特性</p>
                          <p>　　B.14 history对象</p>
                          <p>　　它为window对象的一个属性,该窗口的历史列表。</p>
                          <p>　　属性</p>
                          <p>　　length 历史列表中的项目数</p>
                          <p>　　方法</p>
                          <p>　　back 加载历史列表中的上一个文档</p>
                          <p>　　forward 加载历史列表中的下一个文档</p>
                          <p>　　go 加载历史列表中的一个指定文档,通过文档在列表中的位置来指定</p>
                          <p>　　B.15 image对象</p>
                          <p>　　它是document对象的一个属性,是使用(1MG)标记内嵌入文档里的一幅图像这是Netscape 
                            Navigator 3.0 beta 3引入的新对象。</p>
                          <p>　　属性</p>
                          <p>　　border &lt;IMG&gt;标记的BORDER特性</p>
                          <p>　　complete 表示浏览器是否完整地加载了图像的一个布尔值</p>
                          <p>　　height HEIGHT特性</p>
                          <p>　　hspace HSPACE特性</p>
                          <p>　　lowsrc LOWSRC特性</p>
                          <p>　　src SRC特性</p>
                          <p>　　vsPace VSPACE特性</p>
                          <p>　　width WIDTH特性</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onload</p>
                          <p>　　Onerror</p>
                          <p>　　Onabort</p>
                          <p>　　B.16 images数组</p>
                          <p>　　它是document对象的一个属性,文档中所有图像的列表。</p>
                          <p>　　属性</p>
                          <p>　　length 文档内的图像个数</p>
                          <p>　　B.17 link</p>
                          <p>　　它是document对象的一个属性,文档内的一个&lt;A HREF=&gt;标记。</p>
                          <p>　　属性</p>
                          <p>　　hash 以散列号(#)开始的一个字符串,用于指定URL内的一个锚点</p>
                          <p>　　host 包括冒号和端口号的URL的主机名部分</p>
                          <p>　　hostname 与host属性相同,除了不包括冒号和端口号外</p>
                          <p>　　href 完整的URL</p>
                          <p>　　pathname URL的目录路径部分</p>
                          <p>　　port URL的:端口部分</p>
                          <p>　　protocol URL类型(http:、ftp:、gopher:等等)</p>
                          <p>　　search 以一个问号开始的URL中的一部分,用于指定搜索信息</p>
                          <p>　　target 当用户单击一链接(TARGET特性)时,用于显示被引用文档内容的窗口</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onclick</p>
                          <p>　　Onmouseover</p>
                          <p>　　B.18 links数组</p>
                          <p>　　它是document对象的一个属性,文档内所有链接的一个列表。</p>
                          <p>　　属性</p>
                          <p>　　length 文档内的链接数</p>
                          <p>　　B.19 location对象</p>
                          <p>　　它为document对象的一个属性,该文档的完整URL，请不要把它与window.location属性相混淆,后者可用来加载一个新文档,并且window.location属性本身并不是一个对象，同时,window.location可以用脚本修改,而document.location则不能。</p>
                          <p>　　属性</p>
                          <p>　　hash 以散列号(#)开始的一个字符串,用于指定URL内的一个锚点</p>
                          <p>　　host 包括冒号和端口号的URL的主机名部分</p>
                          <p>　　hostname 与host属性相同,除了不包括冒号和端口号之外</p>
                          <p>　　href 完整的URL</p>
                          <p>　　pathname URL的目录路径部分</p>
                          <p>　　port URL的:端口部分</p>
                          <p>　　protocol URL的类型(http:、ftp:、gopher:等等)</p>
                          <p>　　search 以问号(?)开始的URL中的一部分,用于指定搜索信息</p>
                          <p>　　target 用户单击链接(TARGET特性)时,用于显示被引用文档的内容的窗口</p>
                          <p>　　B.20 math对象</p>
                          <p>　　该对象不是其它对象的一个属性,而是一个内置对象,包含了许多数学常量和函数。</p>
                          <p>　　属性</p>
                          <p>　　E 欧拉常量,自然对数的底(约等于2.718)</p>
                          <p>　　LN2 2的自然对数(约等于0.693)</p>
                          <p>　　LN10 10的自然对数(约等于2.302)</p>
                          <p>　　LOG2E 以2为底的e的对数(约等于1.442)</p>
                          <p>　　LOG10E 以10为底的e的对数(约等于o.434)</p>
                          <p>　　PI ∏的值(约等于3.14159)</p>
                          <p>　　SQRT1_2 0.5的平方根(即l除以2的平方根,约等于o.707)</p>
                          <p>　　SQRT2 2的平方根(约等于1.414)</p>
                          <p>　　方法</p>
                          <p>　　abs() 返回某数的绝对值(即该数与o的距离,例如,2与一2的绝对值都是2)</p>
                          <p>　　acos() 返回某数的反余弦值(以弧度为单位)</p>
                          <p>　　asin() 返回某数的反正弦值(以弧度为单位)</p>
                          <p>　　atan() 返回某数的反正切值(以弧度为单位)</p>
                          <p>　　ceil() 返回与某数相等,或大于该数的最小整数(ceil(-22.22)返回-22;ceil22,22)返回23;ceil(22)返回22)</p>
                          <p>　　cos() 返回某数(以弧度为单位)的余弦值</p>
                          <p>　　exp() 返回en</p>
                          <p>　　floor() 与ceil相反(floor(一22.22)返回一23;floor(22.22)返回22; 
                            floor(22)返回22)</p>
                          <p>　　10g() 返回某数的自然对数(以e为底)</p>
                          <p>　　max() 返回两数间的较大值</p>
                          <p>　　min() 返回两数问的较小值</p>
                          <p>　　pow() 返回m的n次方(其中,m为底,n为指数)</p>
                          <p>　　random() 返回0和1之间的一个伪随机数(该方法仅在Netscape</p>
                          <p>　　Navigator的UNIX版本中有效)</p>
                          <p>　　round() 返回某数四舍五入之后的整数</p>
                          <p>　　sin() 返回某数(以弧度为单位)的正弦值</p>
                          <p>　　sqrt() 返回某数的平方根</p>
                          <p>　　tan() 返回某数的正切值</p>
                          <p>　　B.2l navigator对象</p>
                          <p>　　该对象不是其它对象的属性,而是一个内置对象它包含了有关加载文档的浏览器的信息。</p>
                          <p>　　属性</p>
                          <p>　　appCodeName 浏览器的代码名(例如,Mozilla)</p>
                          <p>　　appName 浏览器的名字</p>
                          <p>　　appVersion 浏览器的版本号</p>
                          <p>　　userAgent 由客户机送到服务器的用户与代理头标文本</p>
                          <p>　　方法</p>
                          <p>　　javaEnabled JavaScript中当前并没有该方法,但是不久之后将会添加上它将查看浏览器是否为兼容JavaScript的浏览器,如果是,继续查看JavaScript是否处于支持状态。</p>
                          <p>　　B.22 options数组</p>
                          <p>　　该数组是select对象的一个属性,即选择框中的所有选项(&lt;OPTION&gt;)的一个列表。</p>
                          <p>　　属性</p>
                          <p>　　defaultSelected 选项列表中的缺省选项</p>
                          <p>　　index 选项列表中某选项的索引位置</p>
                          <p>　　length 选项列表中的选项数(&lt;OPTIONS&gt;)</p>
                          <p>　　name 选项列表的名字(NAME特性)</p>
                          <p>　　selected 表示选项列表中某选项&lt;OPTION&gt;是否被选中的一个布尔类型值</p>
                          <p>　　selectedIndex 选项列表中已选中的&lt;OPTION&gt;的索引(位置)</p>
                          <p>　　text 选项列表中&lt;OPTION&gt;标记后的文本</p>
                          <p>　　value 选项列表中的VALUE=特性</p>
                          <p>　　B.23 Password 对象</p>
                          <p>　　它是document对象的一个属性,一个&lt;INPUT TYPE=”PASSWORD”&gt;标记。</p>
                          <p>　　属性</p>
                          <p>　　defaultValue password对象的缺省值(VAlUE=特性)</p>
                          <p>　　name 对象的名字(NAME=特性)</p>
                          <p>　　value 该域具有的当前值最初与VALUE=特性(defauttValue)相同,但是,如果脚本修改了该域中的值,则该值将改变</p>
                          <p>　　方法</p>
                          <p>　　focus 把焦点从该域移开</p>
                          <p>　　blur 把焦点移到该域</p>
                          <p>　　select 选择输入区域</p>
                          <p>　　B.24 radio对象</p>
                          <p>　　它是form对象的一个属性,窗体内的一组单选按钮(选项按钮)(&lt;INPUT TYPE=”RADIO”&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　checked 复选框或选项按钮(单选按钮)的状态</p>
                          <p>　　defaultChecked 复选框或选项按钮(单选按钮)的缺省状态</p>
                          <p>　　length 一组单选按钮中的按钮数</p>
                          <p>　　name 对象的名字(NAME=特性)</p>
                          <p>　　value VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　click 模拟鼠标单击按钮</p>
                          <p>　　事件处理程序</p>
                          <p>　　onclick</p>
                          <p>　　B.25 reset 对象</p>
                          <p>　　它是form对象的一个属性,复位按钮(&lt;INPUT TYPE=”RESET”&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　name 对象的名字(NAME=特性)</p>
                          <p>　　value VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　click 模拟鼠标单击按钮</p>
                          <p>　　事件处理程序</p>
                          <p>　　onclick</p>
                          <p>　　B.26 select对象</p>
                          <p>　　它是form对象的一个属性,选择框(&lt;SELECT&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　length 选项列表中的选项数(&lt;OPTIONS&gt;)</p>
                          <p>　　name 选项列表的名字(NAME特性)</p>
                          <p>　　options 列表中的选项数</p>
                          <p>　　selectedlndex 选项列表中已选中的&lt;OPTION&gt;的索引(位置)</p>
                          <p>　　text 选项列表中(OPTION)标记之后的文本</p>
                          <p>　　value 选项列表中的VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　blur 把焦点从选项列表中移走</p>
                          <p>　　focus 把焦点移到选项列表中</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onblur</p>
                          <p>　　onchange</p>
                          <p>　　Onfocus</p>
                          <p>　　B.27 string对象</p>
                          <p>　　它不是另一个对象的属性,而是一个内置对象,即一串字符字符串输入脚本中时必须位于引号内。</p>
                          <p>　　属性</p>
                          <p>　　length 字符串中的字符个数</p>
                          <p>　　方法</p>
                          <p>　　anchor() 用来把字符串转换到HTML锚点标记内(&lt;A NAME=&gt;)</p>
                          <p>　　big() 把字符串中的文本变成大字体(&lt;BIG&gt;)</p>
                          <p>　　blink() 把字符串中的文本变成闪烁字体(&lt;BLINK&gt;)</p>
                          <p>　　bold() 把字符串中的文本变成黑字体(&lt;B&gt;)</p>
                          <p>　　charAt() 寻找字符串中指定位置的一个字符</p>
                          <p>　　fixed() 把字符串中的文本变成固定间距字体(&lt;TT&gt;)</p>
                          <p>　　fontcolor() 改变字符串中文本的颜色(&lt;FONT COLOR=&gt;)</p>
                          <p>　　fontsize() 把字符串中的文本变成指定大小(&lt;FONTSIZE=&gt;)</p>
                          <p>　　indexOf() 用来搜索字符串中的某个特殊字符,并返回该字符的索引位置</p>
                          <p>　　italics() 把字符串中的文本变成斜字体(&lt;I&gt;)</p>
                          <p>　　lastlndexOf() 与indexof相似,但是向后搜索最后一个出现的字符</p>
                          <p>　　link() 用来把字符串转换到HTML链接标记中(&lt;A HREF=&gt;)</p>
                          <p>　　small() 把字符串中的文本变成小字体(&lt;SMALL&gt;)</p>
                          <p>　　strike() 把字符串中的文本变成划掉字体(&lt;STRIKE&gt;)</p>
                          <p>　　sub() 把字符串中的文本变成下标(subscript)字体((SUB&gt;)</p>
                          <p>　　substring() 返回字符串里指定位置间的一部分字符串</p>
                          <p>　　sup() 把字符串中的文本变成上标(superscript)字体(&lt;SUP&gt;)</p>
                          <p>　　toLowerCase() 把字符串中的文本变成小写</p>
                          <p>　　toUpperCase() 把字符串中的文本变成大写</p>
                          <p>　　B.28 submit对象</p>
                          <p>　　它是form对象的一个属性,窗体中的一个提交按钮(&lt;INPUT TYPE=”SUBMIT”&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　name 对象的名字(NAME=特性)</p>
                          <p>　　value VALUE=特性</p>
                          <p>　　方法</p>
                          <p>　　click 模拟鼠标单击按钮</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onclick</p>
                          <p>　　B.29 text对象</p>
                          <p>　　它是form对象的一个属性,宙体中的一个文本域(&lt;INPUT TYPE=”TEXT”&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　defaultValue text对象的缺省值(VALUE=特性)</p>
                          <p>　　name 该对象的名字(NAME=特性)</p>
                          <p>　　Value 该域具有的当前值,最初与VALUE=特性(defaultValue)相同但是,如果脚本修改了该域中的值,则该值将改变</p>
                          <p>　　方法</p>
                          <p>　　blur 把焦点从文本框移开</p>
                          <p>　　focus 把焦点移到文本框</p>
                          <p>　　select 选择输入区域</p>
                          <p>　　事件处理程序</p>
                          <p>　　Onblur</p>
                          <p>　　Onchange</p>
                          <p>　　Onfeus</p>
                          <p>　　Onselect</p>
                          <p>　　B.30 textarea对象</p>
                          <p>　　它是form对象的一个属性,宙体中的一个文本区域(&lt;TEXTAREA&gt;)。</p>
                          <p>　　属性</p>
                          <p>　　defaultValue textarea对象的缺省值(VALUE=特性)</p>
                          <p>　　name 该对象的名字(NAME=特性)</p>
                          <p>　　value 该域具有的当前值,最初与VALUE=特性(defaultValue)相同,但是,如果脚本修改了该域中的值,则该值将改变了。</p>
                          <p>　　方法</p>
                          <p>　　blur 把焦点从文本区移开</p>
                          <p>　　focus 把焦点移到文本区</p>
                          <p>　　select 选择输入区域事件处理程序</p>
                          <p>　　Onblur</p>
                          <p>　　Onchange</p>
                          <p>　　Onfocus</p>
                          <p>　　Onselect</p>
                          <p>　　B.31 window对象</p>
                          <p>　　它是一个顶层对象,而不是另一个对象的属性即浏览器的窗口。</p>
                          <p>　　属性</p>
                          <p>　　defaultStatus 缺省的状态条消息</p>
                          <p>　　document 当前显示的文档(该属性本身也是一个对象)</p>
                          <p>　　frame 窗口里的一个框架((FRAME&gt;)(该属性本身也是一个对象)</p>
                          <p>　　frames array 列举窗口的框架对象的数组,按照这些对象在文档中出现的顺序列出(该属性本身也是一个对象)</p>
                          <p>　　history 窗口的历史列表(该属性本身也是一个对象)</p>
                          <p>　　length 窗口内的框架数</p>
                          <p>　　location 窗口所显示文档的完整(绝对)URL(该属性本身也是一个对象)不要把它与如document.location混淆,后者是当前显示文档的URL。用户可以改变window.location(用另一个文档取代当前文档),但却不能改变document.location(因为这是当前显示文档的位置)</p>
                          <p>　　name 窗口打开时,赋予该窗口的名字</p>
                          <p>　　opener 代表使用window.open打开当前窗口的脚本所在的窗口(这是Netscape 
                            Navigator 3.0beta 3所引入的一个新属性)</p>
                          <p>　　parent 包含当前框架的窗口的同义词。frame和window对象的一个属性</p>
                          <p>　　self 当前窗口或框架的同义词</p>
                          <p>　　status 状态条中的消息</p>
                          <p>　　top 包含当前框架的最顶层浏览器窗口的同义词</p>
                          <p>　　window 当前窗口或框架的同义词,与self相同</p>
                          <p>　　方法</p>
                          <p>　　alert() 打开一个Alert消息框</p>
                          <p>　　clearTimeout() 用来终止setTimeout方法的工作</p>
                          <p>　　close() 关闭窗口</p>
                          <p>　　confirm() 打开一个Confirm消息框,用户可以选择OK或Cancel,如果用户单击OK,该方法返回true,单击Cancel返回false</p>
                          <p>　　blur() 把焦点从指定窗口移开(这是Netscape Navigator 3.0 beta 
                            3引入的新方法)</p>
                          <p>　　focus() 把指定的窗口带到前台(另一个新方法)</p>
                          <p>　　open() 打开一个新窗口</p>
                          <p>　　prompt() 打开一个Prompt对话框,用户可向该框键入文本,并把键入的文本返回到脚本</p>
                          <p>　　setTimeout() 等待一段指定的毫秒数时间,然后运行指令事件处理程序</p>
                          <p>　　OnloadOnunload</p>
                          <p> <br>
                          </p></td>
                      </tr>
                    </table> </TD>
                </TR>
              </TBODY>
            </TABLE></td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright2.asp"-->
</body>
</html>