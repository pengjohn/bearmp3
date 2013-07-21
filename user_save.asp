<!--#include file="conn_music_open.asp"-->
<%
dim username
dim useremail
dim useroicq
dim userinfo
dim userimage
dim rs,sql

username = request("username")
useremail = request("useremail")
useroicq = request("useroicq")
userinfo = request("userinfo")
userimage = request("userimage")

set rs=server.CreateObject("ADODB.RecordSet")
rs.open "select * from user where username='"&username&"'",conn,1,3

if (rs.eof and rs.bof) then
	response.write "无此会员"
else
	rs("useremail") = useremail
	rs("useroicq")  = useroicq
	rs("userinfo")  = userinfo
	rs("userimage") = userimage
	rs.update
end if

rs.close
set rs=nothing
%>
<!--#include file="conn_music_close.asp"-->
<%response.redirect "default.asp"%>
