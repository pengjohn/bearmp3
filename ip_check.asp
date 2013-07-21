<%
'''''''###############################################################
'''''''# 精点IP管理器(外挂) 1.0
'''''''###############################################################
dim ipdb,ipconn,stop_ip,un_ip,rs_ip
un_ip=0

'if left(REQUEST.servervariables("REMOTE_ADDR"),10) ="10.100.140" then
'	response.redict "error.htm"
'end if


ipdb=Server.MapPath("ip_forbidden.mdb") ' 在此处更改数据库名称及路径
set ipconn=Server.CreateObject("ADODB.Connection")
ipconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ipdb

 stop_ip= Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If stop_ip="" Then  stop_ip= Request.ServerVariables("REMOTE_ADDR")
call ynip(stop_ip)

if un_ip=0 then 
    set rs_ip=ipconn.execute("select viw From stopip where viw<>0 and ("& cip(stop_ip) &" between oneip and endip)",0,1)
	if not rs_ip.eof then
	        if rs_ip("viw")=1 then Response.write"网站数据维护,请稍后再来<br>"
	        if rs_ip("viw")=2 then Response.write"你的IP被禁止访问<br>"
              rs_ip.close
	          set rs_ip=nothing
			  ipconn.close
              set ipconn=nothing
			Response.end
	end if
	rs_ip.close
	set rs_ip=nothing
end if

sub ynip(uip)
         dim tip
		 tip = split(uip,".") 
		  if not IsArray(tip) then un_ip=1
          if UBound(tip)<3 then un_ip=1
          if Not IsNumeric(tip(0)) or  Not IsNumeric(tip(1)) or  Not IsNumeric(tip(2)) or Not IsNumeric(tip(3)) then un_ip=1
		  if cint(tip(0))>255 or cint(tip(1))>255 or cint(tip(2))>255 or cint(tip(3))>255 then un_ip=1
end sub
function cip(sip)
    dim tip
	'tip=cstr(sip)
	'sip1=left(tip,cint(instr(tip,".")-1))
	'tip=mid(tip,cint(instr(tip,".")+1))
	'sip2=left(tip,cint(instr(tip,".")-1))
	'tip=mid(tip,cint(instr(tip,".")+1))
	'sip3=left(tip,cint(instr(tip,".")-1))
	'sip4=mid(tip,cint(instr(tip,".")+1))
	'if cint(sip1)<128 then
		'cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)
	'else
		'cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)-4294967296
	'end if 
    tip = split(sip,".")
	if cint(tip(0))<128 then
		cip=cint(tip(0))*256*256*256+cint(tip(1))*256*256+cint(tip(2))*256+cint(tip(3))
	else
		cip=cint(tip(0))*256*256*256+cint(tip(1))*256*256+cint(tip(2))*256+cint(tip(3))-4294967296
	end if
end function
ipconn.close
set ipconn=nothing
%>
