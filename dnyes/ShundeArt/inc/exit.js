window.onbeforeunload = function()
{
if (document.body.offsetWidth-50 < event.clientX && event.clientY<0)
return "是否关闭当前窗口？关闭本窗口将自动退出本管理系统！"
}
