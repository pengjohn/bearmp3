<%if Session.Contents("KEY")="super" then%>
	<!--#include file="conn_music_open.asp"-->
	<%conn.execute "delete * from user where userid=" & request("id")
	report_txt = 	"ɾ����Ա���ϳɹ���������ء�"
	%>
	<!--#include file="conn_music_close.asp"-->
<%
else
		report_txt = "�㲻�ǹ���Ա���뷵�ء�"
end if
%>
<script language=vbscript>       
	MsgBox <%=report_txt%>
	location.href = "default.asp?classid=100"
</script> 

