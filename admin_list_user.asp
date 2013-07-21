<!--#include file="conn_music_open.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>骚熊音乐->超级管理员管理</title>
</head>

<body topmargin="0" leftmargin="0">
<div align="center"><center>
<span style="font-size: 9pt">
<table border="1" cellpadding="2" width="750" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF">
	<tr>
		<td width="90%" align="center">
			<table width="100%" height="39" border="1" bordercolor="#FFFFFF">
				<tr>
					<td align="center"><font color="#ffffff">ID号</font></td>
					<td align="center"><font color="#ffffff">用户名</font></td>
					<td align="center"><font color="#ffffff">密码</font></td>
					<td align="center"><font color="#ffffff">权限</font></td>
					<td align="center"><font color="#ffffff">积分</font></td>
					<td align="center"><font color="#ffffff">IP</font></td>
				</tr>
<%
dim rs
set rs=server.CreateObject("ADODB.RecordSet")  
rs.open "select * from user order by userip",conn,1
do while NOT rs.EOF%> 
	<tr>
		<td align="center"><%=rs("userid")%></td>
		<td align="center"><A href='admin_user_login_list.asp?username=<%=rs("username")%>'><%=rs("username")%></a></td>
		<td align="center"><%=rs("userpassword")%></td>
		<td align="center"><%=rs("userrank")%></td>
		<td align="center"><%=rs("userscore")%></td>
		<td align="center"><A href='admin_user_login_list.asp?userip=<%=rs("userip")%>'><%=rs("userip")%></a></td>

	</tr>
<%
rs.MoveNext
loop
rs.close
set rs=nothing
%>
	</table>
    </td>
  </tr>
</table>
</span>
</center></div>
</body>
</html>
<!--#include file="conn_music_close.asp"-->


