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
  errmsg=errmsg+"<li>��ʾ���Ʋ���Ϊ��</li>"
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
<title>ɧ������->�޸ĸ����ɹ�</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" >
<div align="center"><center>
<table border="1" width=80% cellpadding="2" bordercolor="#FFFFFF">
  <tr>
    <td width="100%" height="20"><p align="center"><b>�޸ĳɹ�</b></td>
  </tr>
  <tr>
    <td width="100%"><p align="left"><br>
    ר����ţ�<%=albumid%><br>
    ר�����棺<%=albumcover%><br>
    ר�����ƣ�<%=albumname%></p>
    </td>
  </tr>
</table>
<a href="default.asp">������ҳ</a>
</center></div>
<%else
	response.write "�������µ�ԭ���ܱ������ݣ�"
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>����</a>"
end if
%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
