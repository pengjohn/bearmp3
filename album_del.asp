<%if Session.Contents("IsAdmin")<>true then%>
	<script language=vbscript>       
		MsgBox "�㲻�ǹ���Ա���뷵��"
		location.href = "javascript:history.back()"       
	</script> 
<%end if%>

<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<%
albumid = request("albumid")
conn.execute "delete * from song where albumid=" & albumid
conn.execute "delete * from album where albumid=" & albumid
response.write "ר������ɾ���ɹ���<br>"
%>
<script language=vbscript>       
	MsgBox "ר��ɾ���ɹ����뷵��"
	location.href = "default.asp?singerid=<%=( albumid\(10^Mul_Album) )%>"
</script> 
<!--#include file="conn_music_close.asp"-->
