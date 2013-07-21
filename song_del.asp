<%if Session.Contents("IsAdmin")<>true then%>
	<script language=vbscript>       
		MsgBox "你不是管理员。请返回"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>
<!--#include file="conn_music_open.asp"-->
<%conn.execute "delete * from song where song_id=" & request("song_id")
 response.write "歌曲数据删除成功。<br>"
%>
<script language=vbscript>       
	MsgBox "歌曲删除成功。请返回"
	location.href = "javascript:history.back()"       
</script> 
<!--#include file="conn_music_close.asp"-->
