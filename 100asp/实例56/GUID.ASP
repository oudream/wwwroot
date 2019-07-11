<html>
<head><title>GUID生成器</title></head>
<body>
<blockquote>
<h2 align=center>GUID生成器</h2>
<%
  Set MyGuid = Server.CreateObject ("GuidMakr.GUID")
  Response.Write (MyGuid.GetGUID)
%>
</blockquote>


GUID（读音为 goo-id）即全局唯一标识符，是由确保唯一性的算法生成的一个 128 位（16 字节）的数。该算法来源于开放软件基金会 (OSF) 的分布式计算环境 (DCE)，是关于分布式计算的一套标准。

GUIDs 通常用于唯一地标识 Windows 注册表的入口。例如，Visual Basic 自动生成一个 GUID 来在 Windows 注册表中标识类型库。

Visual Basic 也自动为部件的每一个公共类和接口生成一个 GUID。这些通常称之为类 ID (CLSID) 和接口 ID (IID)。类 ID 和接口 ID 对使用 Visual Basic 开发的部件的版本兼容性是很关键的。

注意   GUIDs 也称为 UUID，即通用唯一标识符。


</body>
</html>