<!--#include file=conn_music_open.asp-->
<!--#include file=conn_function.asp-->
<!--#include file=conn_define.asp-->
<%
dim sql
dim rs
dim username
dim userrank
dim useremail
dim useroicq
dim userscore
dim userinfo
dim userimage

username=request("username")
userrank = "��Ա"

set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&username&"'"
rs.open sql,conn,1,3

if rs.eof and rs.bof then
	response.write "Sorry,û�ҵ��˻�Ա����"
 	response.write "<br> <a href='#'onClick='history.back()'>����</a>"
else
	username=rs("username")
	
	select case rs("userrank")
		case "super"
			userrank = "��������Ա"
		case "check"
			userrank = "�߼�����Ա"
		case "input"
			userrank = "��������Ա"
		case "user"
			userrank = "��ͨ��Ա"
	end select
	
	userscore = rs("userscore")
	useremail = rs("useremail")
	useroicq = rs("useroicq")
	userinfo = rs("userinfo")
	if userinfo <>"" then 
		userinfo=userinfo
	else 
		userinfo="��"	
	end if
		
	userimage = rs("userimage")
%>


<title><%=Info_Title%>->��Ա����</title>
<link rel="stylesheet" type="text/css" href="css/style.css">

<body bgcolor="#008080">
<div align="center"><center>
<table border="0" cellPadding="0" cellSpacing="0">
	<tr>
		<td background="images/card-tl.gif" width="18"></td>
		<td background="images/card-t.gif" height="18"></td>
		<td background="images/card-tr.gif" width="18"></td>
	</tr>
	<tr>
		<td background="images/card-l.gif" width="18"></td>
		<td width="300">
			<table border="0" cellPadding="1" cellSpacing="0" width="95%">
				<tr>
					<td colSpan="3" height="20" style="BORDER-BOTTOM: rgb(0,0,0) 2px solid" width="95%"><font size="3"><%=Info_Title%>��Ա</font></td>
				</tr>
				<tr>
					<td rowSpan="5" style="BORDER-RIGHT: rgb(0,0,0) 1px solid" vAlign="top" width="110"><font size="4"><%=username%></font><br><br>
						<img alt="" height="64" width="64" src="<%=userimage%>">
					</td>
					<td width="45">��ݣ�</td>
					<td><%=userrank%></td>
				</tr>
				<tr>
					<td width="45">���֣�</td>
					<td><%=userscore%></td>
				</tr>
				<tr>
					<td width="45">E_mail��</td>
					<td><a href="mailto:<%=useremail%>"><%=useremail%></a></td>
				</tr>
				<tr>
					<td width="45">OICQ��</td>
					<td><%=useroicq%></td>
				</tr>
				<tr>
					<td width="45" valign="top">ǩ����</td>
					<td height="42"><%=changechr(userinfo)%></td>
      	   </tr>
			</table>
		</td>
		<td background="images/card-r.gif" width="18"></td>
	</tr>
	<tr>
		<td background="images/card-bl.gif" width="18"></td>
		<td background="images/card-b.gif" height="18"></td>
		<td background="images/card-br.gif" width="18"></td>
	</tr>
</table>

</center></div>
</body>
<%
end if
rs.close
set rs=nothing
%>
<!--#include file=conn_music_close.asp-->
