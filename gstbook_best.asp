<!--#include file="conn_music_open.asp"-->
<%
	dim sql 
	dim rs
	dim gstbookid
	gstbookid=request("gstbookid")

	set rs=server.createobject("adodb.recordset")
	sql="select * from gstbook where gstbook_id=" & gstbookid
	rs.open sql,conn,1,3
	if rs.eof then
		str="该帖子不存在！"
	else
		if Session.Contents("IsAdmin")=true then
			if request("mode") = "OK" then
				rs("gstbooktype") = 1
				rs.update
			elseif request("mode") = "NG" then
				rs("gstbooktype") = 0			
				rs.update
			end if			
		else
			str= "对不起！你没有权限标注此贴。"
		end if
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "gstbook.asp"
%>




