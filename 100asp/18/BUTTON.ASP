<html>
<head>
<title>��ֹ����Ҽ�</title>
</head>
<body>
<script language="Javascript">
var mMsg = "�������Ҽ�����ֹ��!";

  function disableRightClick(btnClick)
  {
    if (navigator.appName == "Netscape" && btnClick.which == 3)
    {   
      alert(mMsg);
      return false;
    }
    else if (navigator.appName =="Microsoft Internet Explorer" && event.button == 2)  
    {
      alert(mMsg);
      return false;
    }
  }
document.onmousedown = disableRightClick;
</script>
</body>
</html>
