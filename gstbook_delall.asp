<!--#include file="conn_music_open.asp"-->
<%
	dim sql 
	dim rs

	if Session.Contents("IsAdmin")=true then
			conn.execute "delete * from gstbook"
	else
			str= "对不起！你没有权限!"
	end if
		
	conn.close
	set conn=nothing
	response.redirect "gstbook.asp"
%>
