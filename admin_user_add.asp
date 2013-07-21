<!--#include file="conn_music_open.asp"-->
<%
IF Session.Contents("KEY")="super"THEN
	dim rs,sql
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.open "select * from user where username='"&request("UserName2")&"'",conn,3,2
	if (rs.eof and rs.bof) then
		rs.addnew
			rs("userrank")=request("select")
			rs("username")=request("UserName2")
			rs("userpassword")=request("Passwd2")
			rs("userscore")=100
		rs.update
	end if
	rs.close
	set rs=nothing
end if
%>
<!--#include file="conn_music_close.asp"-->
<%response.redirect "admin_manage.asp"%>