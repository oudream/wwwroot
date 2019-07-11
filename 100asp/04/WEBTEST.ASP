<HTML>
<HEAD>
<TITLE>测试WEB服务器</TITLE>
</HEAD>
<BODY>
<Script language=jscript runat=server>
    response.write ("<table border=1>");
    response.write ("<tr><td>脚本引擎</td><td>");
    response.write (ScriptEngine()+"</td></tr>");
    response.write ("<tr><td>编译版本</td><td>");
    response.write (ScriptEngineBuildVersion()+"</td></tr>");
    response.write ("<tr><td>主 版 本</td><td>");
    response.write (ScriptEngineMajorVersion()+"</td></tr>");
    response.write ("<tr><td>次 版 本</td><td>");
    response.write (ScriptEngineMinorVersion()+"</td></tr>");
    response.write ("</table>");
</script>
<Script language=vbscript runat=server>
   response.write "<table border=1>"
   response.write "<tr><td>脚本引擎</td><td>" 
   response.write ScriptEngine() & "</td></tr>"
   response.write "<tr><td>编译版本</td><td>" 
   response.write ScriptEngineBuildVersion() & "</td></tr>"
   response.write "<tr><td>主 版 本</td><td>" 
   response.write ScriptEngineMajorVersion() & "</td></tr>"
   response.write "<tr><td>次 版 本</td><td>" 
   response.write ScriptEngineMinorVersion() & "</td></tr>"
   response.write "</table>"
</script>
</body>