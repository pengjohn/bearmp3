<!--#include file="conn_music_open.asp"-->
<%
dim song_id , songid, songname, songfile
dim rs, sql

song_id = request("song_id")
songid = request("songid")
songname = request("songname")
songfile = request("songfile")

founerr=false
if trim(request.form("songname"))="" then
	founderr=true
	errmsg=errmsg+"<li>��ʾ���Ʋ���Ϊ��</li>"
end if

if founderr=false then
	set rs=server.createobject("adodb.recordset")
	sql="select * from song where song_id="+cstr(song_id)+""
	rs.open sql,conn,1,3
	rs("songid")  = songid
	rs("songname")= songname
	if request("songfile")<>"" then
		rs("songfile") = songfile
	else		
		rs("songfile") = ""
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

<body bgcolor="#000000">
<div align="center"><center>
<table border="1" width=80% cellpadding="2" bordercolor="#FFFFFF">
	<tr>
		<td width="100%" height="20"><p align="center"><b>�޸ĳɹ�</b></td>
	</tr>
	<tr>
		<td width="100%"><p align="left"><br>
			�������Ϊ��<%=songid%><br>
			�����ļ�Ϊ��<%=songfile%><br>
			��������Ϊ��<%=songname%></p>
		</td>
	</tr>
</table>
</center></div>
<%else
	response.write "�������µ�ԭ���ܱ������ݣ�"
	response.write errmsg
	response.write "<br> <a href='#'onClick='history.back()'>����</a>"
end if%>
<br><br>
</body>
</html>
<!--#include file="conn_music_close.asp"-->
