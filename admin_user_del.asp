<%if Session.Contents("KEY")="super" then%>
	<!--#include file="conn_music_open.asp"-->
	<%conn.execute "delete * from user where userid=" & request("id")
	report_txt = 	"删除会员资料成功！点击返回。"
	%>
	<!--#include file="conn_music_close.asp"-->
<%
else
		report_txt = "你不是管理员！请返回。"
end if
%>
<script language=vbscript>       
	MsgBox <%=report_txt%>
	location.href = "default.asp?classid=100"
</script> 

