<!--#include file="conn_music_open.asp"-->
<%
dim album_id , albumid, albumname, albumcover, albumgrade
dim rs, sql

album_id = request("album_id")
albumid = request("albumid")
albumname = request("albumname")
albumcover = request("albumcover")
albumgrade = request("albumgrade")

founerr=false
if trim(request.form("albumname"))="" then
  founderr=true
  errmsg=errmsg+"<li>显示名称不能为空</li>"
end if

if founderr=false then
	set rs=server.createobject("adodb.recordset")
	sql="select * from album where album_id="+cstr(album_id)+""
	rs.open sql,conn,1,3
	rs("albumid")  = albumid
	rs("albumname")= albumname
	rs("albumgrade")= albumgrade
	if request("albumcover")<>"" then
		rs("albumcover") = albumcover
	else		
		rs("albumcover") = "images/nocover.gif"
	end if
	rs.update
	rs.close
	set rs=nothing
%>
<html>
<head>
<title>骚熊音乐->修改歌曲成功</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<div align="center"><center>
<table border="1" width=80% cellpadding="2" bordercolor="#FFFFFF">
  <tr>
    <td width="100%" height="20"><p align="center"><b>修改成功</b></td>
  </tr>
  <tr>
    <td width="100%"><p align="left"><br>
    专辑编号：<%=albumid%><br>
    专辑封面：<%=albumcover%><br>
    专辑名称：<%=albumname%></p>
    </td>
  </tr>
</table>
<a href="default.asp">返回首页</a>
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
