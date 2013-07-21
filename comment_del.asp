<%
if Session.Contents("IsAdmin")=true then
   dim comment_id
   dim commentindex
   comment_id=request("comment_id")
   commentindex=request("commentindex")
%>
	<!--#include file="conn_music_open.asp"-->
   <%conn.execute "delete * from comment where comment_id=" & comment_id%>
	<!--#include file="conn_music_close.asp"-->
	<%response.redirect "comment_show.asp?commentindex="&commentindex
else
   response.Write "对不起！请先登陆成为系统管理员。"
end if
%>


