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
 		response.write "注册成功！"
%>
	<script language=vbscript>
		MsgBox "注册成功！请从首页登陆。"
		location.href = "default.asp"
	</script> 
<%	else
	 	response.write "用户名"
		response.write request("username")
	 	response.write "已经存在！"
 		response.write "<br> <a href='user_reg.asp'>请返回重新注册</a>"
	end if
	rst.close
	set rst=nothing
else
 	response.write "用户名或密码不能为空！"
 	response.write "<br> <a href='user_reg.asp'>请返回重新注册</a>"
end if
%>
<!--#include file="conn_music_close.asp"-->

