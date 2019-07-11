<%
Function ComputerChoose()
    Dim RandomNum
    Dim choice
    Randomize
    RandomNum=int(rnd*3)+1
    if RandomNum=1 then
       choice="R"
    elseif RandomNum=2 then
       choice="S"
    else
       choice="P"
    end if
    ComputerChoose=choice
End Function

Sub DetermineWinner(playerChoice,computerChoice)
    if playerChoice="R" then
       if computerChoice="R" then
          response.write "<font color=red>石头</font>-----<font color=blue>石头</font><br>"
          response.write "我们平手了"
       elseif computerChoice="S" then
          response.write "<font color=red>石头</font>-----<font color=blue>剪子</font><br>"
          response.write "恭喜你，你赢了"
       elseif computerChoice="P" then
          response.write "<font color=red>石头</font>-----<font color=blue>布</font><br>"
          response.write "恭喜你，你赢了"
       end if
    elseif playerChoice="S" then
       if computerChoice="R" then
          response.write "<font color=red>剪子</font>-----<font color=blue>石头</font><br>"
          response.write "恭喜你，你赢了"
       elseif computerChoice="S" then
          response.write "<font color=red>剪子</font>-----<font color=blue>剪子</font><br>"
          response.write "我们平手了"
       elseif computerChoice="P" then
          response.write "<font color=red>剪子</font>-----<font color=blue>布</font><br>"  
          response.write "恭喜你，你赢了"
       end if
    elseif playerChoice="P" then
       if computerChoice="R" then
          response.write "<font color=red>布</font>-----<font color=blue>石头</font><br>"
          response.write "恭喜你，你赢了"
       elseif computerChoice="S" then
          response.write "<font color=red>布</font>-----<font color=blue>剪子</font><br>"
          response.write "这次我赢了"
       elseif computerChoice="P" then
          response.write "<font color=red>布</font>-----<font color=blue>布</font><br>"
          response.write "我们平手了"
       end if
    end if
End sub
%>
<html>
<title>石头、剪子、布小游戏</title>
<body>
<center>
玩家:<font color=white>------</font>电脑:<br>
<%
dim player
dim computer
player=request("choice")
if player="P" or player="R" or player="S" then
    computer=ComputerChoose()
    call DetermineWinner(player,computer)
end if
%>
<hr>
请选择你的武器:
<form action=index.asp method=post>
<input type=radio name=choice value="R">石头<br>
<input type=radio name=choice value="S">剪子<br>
<input type=radio name=choice value="P">布 &nbsp;<br>
<input type=submit value=开战>
</form>
</center>
</body>
</html>
