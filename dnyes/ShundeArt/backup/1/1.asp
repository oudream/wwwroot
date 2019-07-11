<title>调用演示</title>

调用参数测试:<br>
[shu]显示数目[order]排序方式,click为按点击率,[title]标题显示字数,[click]是否显示点击数,[time]是否显示时间.<br>
写法:<br>
[script src=news.asp?shu=8&order=click&title=3&time=0&click=1]<br>

<script src=news.asp?shu=8&order=click&title=3&time=0&click=1></script><br>
<title>调用演示</title>
调用总栏ID号为11的文章
<script src=news.asp?typeid=11></script><br>
调用大类ID号为15的文章
<script src=news.asp?bigclassid=15&title=10></script><br>
调用小类ID号为46的文章
<script src=news.asp?smallclassid=46></script><br>
调用专栏ID号为16的文章
<script src=news.asp?specialid=16></script><br>
调用总栏ID号为10的图片文章
<script src=picnews.asp?typeid=10&shu=5></script><br>
调用大类ID号为29的图片文章
<script src=picnews.asp?bigclassid=29></script><br>
调用小类ID号为52的图片文章
<script src=picnews.asp?smallclassid=52&title=20></script><br>
调用专栏ID号为16的图片文章
<script src=picnews.asp?specialid=16></script>
调用留言和文章评论[作者: A09，演示: http://heiyan.yeah.net]
<!--#include file=reviewindex.asp-->

调用最新公告
<!--#include file=boardnews.asp-->
