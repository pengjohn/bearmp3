<!--#include file="count_conn.asp"-->
<%
set rs_counter=server.createobject("adodb.recordset")
set rs_login=server.createobject("adodb.recordset")
rs_counter.Open "Select * From counter order by ID desc" ,conn_count,1,3

lastip = rs_counter("lastip")
total  = rs_counter("total")
newip = REQUEST.servervariables("REMOTE_ADDR") 

if cstr(day(rs_counter("date"))) <> cstr(day(date())) then
	rs_counter.addnew
	rs_counter("date")  = date()                      
	rs_counter("today") = 1
	rs_counter("total") = total + 1
	rs_counter("lastip") =  newip
	rs_counter.update
	Session.Contents("IsFirst")="No"

	rs_login.Open "Select * From login where (loginid is null)" ,conn_count,1,3
	rs_login.addnew
	rs_login("login_ip")   = newip
	if Session.Contents("mp3_username")<>"" then
		rs_login("login_name") = Session.Contents("mp3_username")
	else
		rs_login("login_name") = "����"
	end if
	rs_login("login_time") = now()
	rs_login.update
	rs_login.close
else
	if Session.Contents("IsFirst")="" then
		Session.Contents("IsFirst")="No"
		rs_login.Open "Select * From login where login_time>date() and login_ip='"&cstr(newip)&"'" ,conn_count,1,3
		if not rs_login.eof then
'pengjohn add,ͳ�����з���
			rs_login.addnew
			rs_login("login_ip")   = newip
			rs_login("login_name") = "����"
			rs_login("login_time") = now()
			rs_login.update
			rs_login.close
'---------------------------
		else
			rs_counter("total")  =  rs_counter("total") + 1               
			rs_counter("today")  =  rs_counter("today") + 1               
			rs_counter("lastip") =  newip
			rs_counter.update                                    

			rs_login.addnew
			rs_login("login_ip")   = newip
			rs_login("login_name") = "����"
			rs_login("login_time") = now()
			rs_login.update
			rs_login.close
		end if
	end if
end if

N = Now
D1 =  #8/23/2001#               ' ��ʼͳ������(��/��/��)
D2 = DateValue(N)
D3 = DateDiff("d", D1, D2)	' ����ʱ��-��ʼʱ��

response.write "�� <b>"&D1&"</b> ��վ����"
response.write "<br>����������<b>"&rs_counter("total")&"</b>"
response.write "<br>����������<b>"&rs_counter("today")&"</b>"

rs_counter.Close
set rs_login=nothing
conn_count.close
set conn_count=nothing
%>
