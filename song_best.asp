<!--#include file="conn_music_open.asp"-->
<%
dim rs,sql
dim song_id
dim best

song_id = request("song_id")

set rs=server.CreateObject("ADODB.RecordSet")
sql = "select * from song where song_id="&song_id
rs.open sql,conn,1,3

if (rs.eof and rs.bof) then
	response.write "����������"
else
	if rs("songhot")=5 then
		rs("songhot") = 1
		best = False
	else
		rs("songhot") = 5
		best = True
	end if
	rs.update
end if
rs.close
set rs=nothing
%>
<!--#include file="conn_music_close.asp"-->

<%if best=True then
	response.write "�Ƽ������ɹ���<br>"
else
	response.write "ȡ�������Ƽ��ɹ���<br>"
end if
%>
<script language=vbscript>       
	MsgBox "���أ�"
	location.href = "javascript:history.back()"       
</script> 
