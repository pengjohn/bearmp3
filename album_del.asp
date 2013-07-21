<%if Session.Contents("IsAdmin")<>true then%>
	<script language=vbscript>       
		MsgBox "你不是管理员。请返回"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%
albumid = request("albumid")
conn.execute "delete * from song where albumid=" & albumid
conn.execute "delete * from album where albumid=" & albumid
response.write "专辑数据删除成功。<br>"
%>
<script language=vbscript>       
	MsgBox "专辑删除成功。请返回"
	location.href = "default.asp?singerid=<%=( albumid\(10^Mul_Album) )%>"
</script> 
<!--#include file="conn_music_close.asp"-->
