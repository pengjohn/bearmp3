<%
dim sql
dim rs
dim seekerrs
dim founduser
dim username
dim companyid
dim password
dim errmsg
dim founderr
dim loginmode

username=Request.form("username")
password=Request.form("password")
loginmode=request("loginmode")


if username="" or password="" then%>
	<script language=vbscript>       
		MsgBox "会员名或密码不能为空！"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="count_conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from user where username='"&username&"'and userpassword='"&password&"'"
rs.open sql,conn,2,3
if not rs.EOF then
	Session.Contents("mp3_username")=rs("username")
	Session.Contents("KEY")=rs("userrank")
'	if Session.Contents("KEY")="check" or Session.Contents("KEY")="super" then
	if Session.Contents("KEY")="super" then
		Session.Contents("IsAdmin")=true
	else
   	Session.Contents("IsAdmin")=false
	end if

	set rslogin=server.createobject("adodb.recordset")
	rslogin.Open "Select * From login where (loginid is null)" ,conn_count,1,3
		rslogin.addnew
		rslogin("login_ip")   = REQUEST.servervariables("REMOTE_ADDR")
		rslogin("login_name") = Session.Contents("mp3_username")
		rslogin("login_time") = now()
		rslogin.update
	rslogin.close
	set rslogin=nothing

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing		
	
if loginmode = 1 then
		response.redirect "default_qvga.asp"
	else
	  response.redirect "default.asp"
	end if
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing		
	response.write "Sorry,会员名或密码错误！"
%>	
	<script language=vbscript>       
		MsgBox "请输入正确的会员名和密码"       
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>
