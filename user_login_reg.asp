<!--#include file="conn_music_open.asp"-->
<%
dim rs,tsql
dim rst
username = request("username")
password = request("password")

if request("password")<>"" and request("username")<>"" then
	set rst=server.CreateObject("ADODB.RecordSet")
	rst.open "select * from user where username='"&request("username")&"'",conn,3,2
	if (rst.eof and rst.bof) then
		rst.addnew
		rst("userrank") = "user"
		rst("username") = username
		rst("userpassword") = password
		rst("userscore") = 100
		rst("userimage") = "pic\Image001.gif"
		rst("userip") = Request.ServerVariables("REMOTE_ADDR")
		rst.update
		response.write username
 		response.write "ע��ɹ���"
%>
	<script language=vbscript>
		MsgBox "ע��ɹ��������ҳ��½��"
		location.href = "default.asp"
	</script> 
<%	else
	 	response.write "�û���"
		response.write request("username")
	 	response.write "�Ѿ����ڣ�"
 		response.write "<br> <a href='user_reg.asp'>�뷵������ע��</a>"
	end if
	rst.close
	set rst=nothing
else
 	response.write "�û��������벻��Ϊ�գ�"
 	response.write "<br> <a href='user_reg.asp'>�뷵������ע��</a>"
end if
%>
<!--#include file="conn_music_close.asp"-->

