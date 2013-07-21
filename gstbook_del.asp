<!--#include file="conn_music_open.asp"-->
<%
	dim sql 
	dim rs
	dim gstbookid
	gstbookid=request("gstbookid")

	set rs=server.createobject("adodb.recordset")
	sql="select * from gstbook where gstbook_id=" & gstbookid
	rs.open sql,conn,1,1
	if rs.eof then
		str="该帖子不存在！"
	else
		if Session.Contents("IsAdmin")=true or Session.Contents("KEY")="check" or Session.Contents("mp3_username")=rs("gstbookname") then
			if rs("gstbooklayer")>1 then
				conn.execute "delete * from gstbook where gstbook_id=" & gstbookid
			else
				conn.execute "delete * from gstbook where gstbookTopID=" & rs("gstbookTopID")
			end if
			str= "删除成功。"
		else
			str= "对不起！你没有权限删除此贴。"
		end if
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "gstbook.asp"
%>
