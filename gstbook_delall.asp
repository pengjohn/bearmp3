<!--#include file="conn_music_open.asp"-->
<%
	dim sql 
	dim rs

	if Session.Contents("IsAdmin")=true then
			conn.execute "delete * from gstbook"
	else
			str= "�Բ�����û��Ȩ��!"
	end if
		
	conn.close
	set conn=nothing
	response.redirect "gstbook.asp"
%>
