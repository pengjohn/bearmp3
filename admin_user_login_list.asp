<!--#include file="count_conn.asp"-->
<span style="font-size: 9pt">
<%
set rs=server.createobject("adodb.recordset")
username = request("username")
userip = request("userip")

if username <>"" then
	sql = "Select * From login where login_name='"&username&"' order by login_time desc"
	rs.open sql,conn_count,0,1
	do while not rs.eof
		response.write rs("login_name")&"["&rs("login_ip")&"]["&rs("login_time")&"]<br>"
	rs.movenext   
	loop
	rs.close
else
	if userip <>"" then
		sql = "Select * From login where login_ip='"&userip&"' order by login_time desc"
		rs.open sql,conn_count,0,1
		do while not rs.eof
			response.write rs("login_name")&"["&rs("login_ip")&"]["&rs("login_time")&"]<br>"
		rs.movenext   
		loop
		rs.close
	else
		rs.Open "Select * From login where DateDiff('d',DateValue(login_time),DateValue(Now))<7 order by login_time desc" ,conn_count,0,1
		do while not rs.eof
			response.write rs("login_name")&"["&rs("login_ip")&"]["&rs("login_time")&"]<br>"
		rs.movenext   
		loop
		rs.close
	end if
end if

set rs=nothing
conn_count.close
set conn_count=nothing
%>
</span>