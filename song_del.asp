<%if Session.Contents("IsAdmin")<>true then%>
	<script language=vbscript>       
		MsgBox "�㲻�ǹ���Ա���뷵��"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>
<!--#include file="conn_music_open.asp"-->
<%conn.execute "delete * from song where song_id=" & request("song_id")
 response.write "��������ɾ���ɹ���<br>"
%>
<script language=vbscript>       
	MsgBox "����ɾ���ɹ����뷵��"
	location.href = "javascript:history.back()"       
</script> 
<!--#include file="conn_music_close.asp"-->
