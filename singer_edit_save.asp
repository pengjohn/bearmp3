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
  errmsg=errmsg+"<li>�������Ʋ���Ϊ��</li>"
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
<title>ɧ������->�޸ĸ��ֳɹ�</title>
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
		���ֱ�ţ�<%=singerid%><br>
		�������ƣ�<%=singername%></p>
		</td>
	</tr>
</table>
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
