<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 

'''''''###############################################################
'''''''# ����IP������(���) 1.0
'''''''###############################################################

'''''''################## ��֤����Ա��� #########################
dim webname,pw
pw=request("pw")
if pw="admin" then Session("stopip")="1"
if Request.QueryString("cls")="exit" then Session("stopip")=""

'''''''###############################################################
if Session("stopip")<>"" then
dim ipdb,ipconn,ipid,ipcls
ipdb=Server.MapPath("ip_forbidden.mdb") ' �ڴ˴��������ݿ����Ƽ�·��
set ipconn=Server.CreateObject("ADODB.Connection")
ipconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ipdb

ipid=Request.QueryString("id")
ipcls=Request.QueryString("cls")


Select case ipcls
case ""

case "del"
	ipconn.execute("Delete  From stopip Where id="&ipid)
case "jd0"
    ipconn.execute("Update stopip set viw=0 Where id="&ipid)
case "jd1"
    ipconn.execute("Update stopip set viw=1 Where id="&ipid)
case "jd2"
     ipconn.execute("Update stopip set viw=2 Where id="&ipid)
end Select

if Request.Form("Submit")="�ύ" then
dim oneip,endip
oneip=trim(Request.Form("oneip"))
endip=trim(Request.Form("endip"))
call ynip(oneip)
call ynip(endip)
    Set rs = Server.CreateObject("ADODB.RecordSet")
	sql="Select * from stopip"
	Rs.open sql,ipconn,2,3
	Rs.addnew
	      rs("oneip")=cip(oneip)
		  rs("endip")=cip(endip)
		  rs("ip1")=oneip
		  rs("ip2")=endip
		  rs("rdate")=now()
		  rs("viw")=1
	rs.update
	rs.close
end if
end if 
 stop_ip= Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If  stop_ip= "" Then  stop_ip= Request.ServerVariables("REMOTE_ADDR")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����IP������</title>
<style type="text/css">
<!--
.unnamed1 {
	font-size: 12px;
	line-height: 20px;
}
-->
</style>
</head>

<body>
<% If  Session("stopip")="" Then %>
<form name="form1" method="post" action="ip_admin.asp">
  <TABLE width="600" height=0 border=1 align=center cellPadding=3 cellSpacing=0 class="unnamed1" style="border-collapse:collapse" >
    <tr bgcolor="#E4F2FA"> 
      <td colspan="3"> 
        <div align="center">��¼ <%=stop_ip%></div></td>
    </tr>
    <tr> 
      <td width="19%"><div align="right"></div></td>
      <td width="30%"><input name="pw" type="password" id="pw"></td>
      <td width="51%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><input type="submit" name="Submit1" value="�ύ"></td>
    </tr>
  </table>
  <br>
</form>
<%
Response.End
 end if %>

<form name="form1" method="post" action="">
  <TABLE width="600" height=0 border=1 align=center cellPadding=3 cellSpacing=0 class="unnamed1" style="border-collapse:collapse" >
    <tr bgcolor="#E4F2FA"> 
      <td colspan="3"> 
        <div align="center">����һ��IP���� <%=stop_ip%></div></td>
    </tr>
    <tr> 
      <td width="19%"><div align="right">IP��:</div></td>
      <td width="30%"><input name="oneip" type="text" id="oneip"></td>
      <td width="51%"><input name="endip" type="text" id="endip"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><input type="submit" name="Submit" value="�ύ"></td>
    </tr>
  </table>
  <br>
</form>
<% set rs=ipconn.execute("select * from stopip")%>
<TABLE width="600" height=0 border=1 align=center cellPadding=0 cellSpacing=0 class="unnamed1" style="border-collapse:collapse" >
  <TBODY>
    <TR > 
      <TD  height=25 bgcolor="#E4F2FA" class="jd_title" > 
        <div align="center">����IP��Ϣ</div></TD>
    </TR>
    <TR> 
      <TD  align=middle valign="top" > <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0"  class="unnamed1">
          <tr>
            <td>״̬</td>
            <td>��ʼ</td>
            <td>��ֹ</td>
            <td>ʱ��</td>
            <td>����</td>
          </tr>
          <tr> 
            <td height="2" colspan="5" bgcolor="#E4F2FA" class="jd_title"></td>
          </tr>
          <%			ipi=0
do while not rs.eof 
ipi=ipi+1 %>
          <tr>
            <td>
              <% 
			  if rs("viw")=1 then
			  Response.write("<font color='#FF9900'>")
			  end if
			  if rs("viw")=0 then
			  Response.write("<font color='#009900'>")
			  end if
			   if rs("viw")=2 then
			  Response.write("<font color='#FF0000'>")
			  end if
              Response.write("��</font>")
			  %>
            </td>
            <td><%=rs("ip1")%></td>
            <td><%=rs("ip2")%></td>
            <td><%=rs("rdate")%></td>
            <td><a href="ip_admin.asp?cls=del&id=<%=rs("id")%>" onClick="return confirm('ɾ��һ��IP�Σ���\n\n�ò������ɻָ���������Ϊ��ͨ,���������.\n\nҪɾ���밴[ȷ��]\n\n��С�İ���[ȡ��]���أ�')">ɾ��</a> 
              <a href="ip_admin.asp?cls=jd0&id=<%=rs("id")%>">��ͨ</a> 
              <a href="ip_admin.asp?cls=jd1&id=<%=rs("id")%>">ά��</a> 
              <a href="ip_admin.asp?cls=jd2&id=<%=rs("id")%>">��ֹ</a> 
            </td>
          </tr>
          <tr> 
            <td height="1" class="jd_title" colspan="5" bgcolor="#E4F2FA"></td>
          </tr>
          <% 
		rs.movenext
    	loop
		rs.close
		set rs=nothing
		ipconn.close
		set ipconn=nothing
%>
          <tr> 
            <td colspan="5">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="5"><div align="center"><font color="#009900">��ͨ </font><font color="#FF9900">��ά����</font> 
                <font color="#FF0000"> ���</font></div></td>
          </tr>
        </table>
</TD>
    </TR>
  </TBODY>
</TABLE>
<div align="center"><br>
  <a href="ip_admin.asp?cls=exit">�˳�</a><br>
</div>
</body>
</html>
<% ' IP�㷨
function cip(sip)
	tip=cstr(sip)
	sip1=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip2=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip3=left(tip,cint(instr(tip,".")-1))
	sip4=mid(tip,cint(instr(tip,".")+1))
	if cint(sip1)<128 then
		cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)
	else
		cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)-4294967296
	end if
end function

sub eorr()
			         Response.Write "<script language=javascript>alert(""��������ȷ��IP"");location.href=""ip_admin.asp"";</script>"
		             Response.End
end sub
sub ynip(uip)
         dim tip
		 tip = split(uip,".") 
		  if not IsArray(tip) then call eorr()
          if UBound(tip)<3 then call eorr()
          if Not IsNumeric(tip(0)) or  Not IsNumeric(tip(1)) or  Not IsNumeric(tip(2)) or Not IsNumeric(tip(3)) then call eorr()
		  if cint(tip(0))>255 or cint(tip(1))>255 or cint(tip(2))>255 or cint(tip(3))>255 then call eorr()
end sub
%>