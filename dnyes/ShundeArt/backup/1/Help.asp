<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 审核评论</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
  <table width="100%" border="0" cellspacing="1" cellpadding="3">
    <tr> 
      <td height="25"  bgcolor="#0099FF"><font color="#FFFFFF"><strong>使用说明：</strong></font></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC">1、如果出现了红色的“对不起，你没有执行该操作的权限！”的字样时，您可以重新登录试试！如果还出现这样的情况，说明您无权操作该部分功能！<br> 
        <br>
        2、用户第一次登录时必须修改自己的资料，除了自我介绍为可选项外，其它必须填，尤其是用户的密码。<br> <br>
        3、在用户生日那天，当该用户登录系统时，系统会对该用户的生日表示祝贺。<br> <br>
        4、该系统目前支持这几类用户：超级管理员、系统管理员、文章审核员、总栏管理员、大类管理员、小类管理员及注册用户。<br> <br>
        5、超级管理员可以管理该系统的全部功能；系统管理员可管理该系统的绝大部分功能；文章审核员可以在所有的类别中添加、修改及审核文章；总栏管理员可管理本总栏的全部类别和文章；大类管理员可管理本大类的全部类别和文章；小类管理员可管理本小类自己添加的文章；注册用户可在任意类别添加文章。<br> 
        <br>
        6、对于上传图片和附件，本系统采用了两种上传方式：有组件和无组件，用户可以根据自己的情况自由选择。如果有条件，建议采用有组件方式，注册方法：双击 
        dll 目录中的“组件注册.bat”来注册组件。<br> <br>
        7、邮件发送采用系统自带的 CDONTS，目前不支持多种邮件发送组件选择。<br> <br>
        8、本系统采用了所见即所得的仿 Word 界面文章编辑器，可实现图文混排，并对 Word 文档提供了绝佳的支持，用户可从 Word 文档中直接复制内容到该编辑器中进行编辑。<br> 
        <br>
        9、本系统支持的图片格式为：jpg、gif、png、bmp；支持的附件格式为：swf、doc、xls、ppt、zip、rar、asf、avi、wmv、rm；图片默认上传尺寸为 
        200K，如需修改默认上传尺寸，可修改系统 admin 目录下的文件 upfile.asp 中的数字“204000”，将之改成适合的尺寸（单位为字节）。<br> 
        <br>
        10、本系统在 Windows 2000+SP3、Access 2000、IE 5.5环境下测试通过；推荐使用 Windows 2000。</td>
    </tr>
    <tr> 
      <td bgcolor="#0099FF"><strong><font color="#FFFFFF">你问我答：</font></strong></td>
    </tr>
    <tr>
      <td height="898" bgcolor="#CCCCCC"> 
        <p>你问我答： <br>
          问：系统数据上的访问总数好象是还包含以前的, 能不能在那里可以改啊?<br>
          答：INCLUDE/CHY.CNT</p>
        <p>问：我的会员排行榜只能显示一个人？<br>
          答：只有发过文的会员才会在排行榜上显示……</p>
        <p>问：最新图文这块怎么弄啊？ 我写一篇文章，并有图片， 怎么在主页上没有显示啊？ 只是在文章中显示？ 这是什么原因啊？<br>
          答：你在编辑文章的时候，如果您上传了图片，请在后面单选框中选择“是”</p>
        <p>问：我怎么觉得下载模板和新闻模板没什么区别呢？<br>
          答：实际并没有设不同的模板，都是用同一个，当然可以设置不同的显示形式</p>
        <p>问：我怎么无法修改发表过的文章了？<br>
          答：每篇文章后面有按钮，如果没有在“大类”下面建立“小类”的话，点击“本大类无小类的文章”按钮，就可以看到文章了。</p>
        <p>问：为什么发送邮件给好友功能没有用处？<br>
          答：需要服务器支持</p>
        <p>问：首页“论坛最新”论坛连接到我的论坛了，可怎么不更新呢？<br>
          答：修改index.asp，论坛需要下载动网论坛插件－首页调用</p>
        <p>问：今日导读这块怎么删除啊？<br>
          答：用dreamweaver或其他编辑器打开index.asp,直接del</p>
        <p>问：“本站专题”的实际作用是什么？<br>
          答：可以是在不同的栏目中，却又与某一相同问题有关的一系列文章的集合</p>
        <p>问：超级管理员的用户名和密码可不可以改的啊？<br>
          答：系统管理-&gt;用户管理里改</p>
        <p>问：在win2000下总是出现“网页无法显示，访问的人数太多”，尘缘系统是不是有人数访问限制啊？<br>
          答：出现这种情况是因为你的操作系统安装的IIS的访问人数限制,跟尘缘系统无关!建议安装win2000server或者以上版本!</p>
        <p>问：在后台，单位类型怎样修改添加？<br>
          答：看看数据库</p>
        <p>问：请问新添加的广告功能在首页的那里显示？<br>
          答：你想在哪里显示都可以：您在管理页面添加广告条目后，到需要放置广告的页面的指定位置添加如下代码： </p>
        <p>网站首页、栏目首页或黄金广告位置的代码为： <br>
          &lt;script language=javascript src=http://网站/广告目录/ads.asp?place=1&gt;&lt;/script&gt;</p>
        <p>网站一般内容页面或者非黄金广告位置的代码为： <br>
          &lt;script language=javascript src=http://网站/广告目录/ads.asp?place=0&gt;&lt;/script&gt;</p>
        <p>问：我想问一下，如果要最新图文在首叶显示多个图片，要做怎样的修改啊！<br>
          答：后台系统管理-&gt;网站属性</p>
        <p>问：0。32版在完成系统初始化后，首页出错!为什么？？？<br>
          答：必需得从新再设置一个总栏，就可以了！</p>
        <p>问：在首页中，每个大栏目的文章数目怎么不定呀<br>
          答：后台系统管理-&gt;网站属性</p>
        <p>问：把管理员登陆的那一项给屏蔽了，有什么办法进入管理界面吗？<br>
          答：admin/login.asp</p>
        <p><br>
          问：我一上传图片就服务器对象错误<br>
          答：使用无组件上传模式试试看，在系统设置里改</p>
        <p>问：如何使计数器清零？<br>
          答：修改include\chy.cnt-------------访问总数<br>
          修改include\dchy.cnt-----------今日访问数（这个好像没什么必要）</p>
        </td>
    </tr>
  </table>
</div>
<!--#include file=Admin_Bottom.asp-->
</BODY>
</HTML>