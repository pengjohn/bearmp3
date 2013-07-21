<!--#include file="conn_music_open.asp"-->
<%
dim singerid , singername, singerindex
dim rs, sql

singerid  = request("singerid")
singername= request("singername")
singerindex= request("singerindex")

founerr=false
if trim(request.form("singername"))="" then
  founderr=true
  errmsg=errmsg+"<li>歌手名称不能为空</li>"
end if

if founderr=false then
    set rs=server.createobject("adodb.recordset")
    sql="select * from singer where singerid="+cstr(singerid)+""
    rs.open sql,conn,1,3
		rs("singername")= singername
		rs("singerindex")= singerindex
    rs.update
    rs.close
    set rs=nothing
%>
<html>
<head>
<title>骚熊音乐->修改歌手成功</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body bgcolor="#000000">
<div align="center"><center>
<table border="1" width=80% cellpadding="2" bordercolor="#FFFFFF">
	<tr>
		<td width="100%" height="20"><p align="center"><b>修改成功</b></td>
	</tr>
	<tr>
		<td width="100%"><p align="left"><br>
		歌手编号：<%=singerid%><br>
		歌手名称：<%=singername%></p>
		</td>
	</tr>
</table>
</center></div>
<%else
	response.write "由于以下的原因不能保存数据："
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>返回</a>"
end if
%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
