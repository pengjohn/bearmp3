<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<!--#include file="conn_function.asp"-->

<%
Session.Timeout=600
set rs=server.createobject("adodb.recordset")
Admin=Session.Contents("IsAdmin")

if not isempty(request("page")) then
	currentPage=cint(request("page"))
	thispage=currentpage
else
	currentPage=1
	thispage=currentpage
end if

cmd = request("cmd")
%>

<HTML>
<HEAD>
<TITLE><%=Info_Title%><%=Info_Version%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">

<SCRIPT language=JavaScript>
function MM_showHideLayers()
	{
	var i,visStr,args,theObj;
	args=MM_showHideLayers.arguments;
	for(i=0;i<(args.length-2);i+=3)
		{
		visStr=args[i+2];
		if(navigator.appName=='Netscape'&&document.layers!=null)
			{
			theObj=eval(args[i]);
			if(theObj)
				theObj.visibility=visStr;
			}
		else if(document.all!=null)
			{
			if(visStr=='show')
				visStr='visible';
			if(visStr=='hide')
				visStr='hidden';
			theObj=eval(args[i+1]);
			if(theObj)
				theObj.style.visibility=visStr;
			}
		}
	}

function openScript(url, width, height)
	{
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=no,menubar=no,status=yes' );
	}
function playonline(url, width, height)
	{
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=no,menubar=no,status=no' );
	}
</script>
</HEAD>

<BODY bgColor=#000000 leftMargin=0 topMargin=0 align="center">
<DIV align=center><CENTER>
<TABLE style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width=776 border=0>
	<TR>
		<TD><IMG height=1 src="Images/spacer.gif" width=108></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=1></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=50></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=36></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=18></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=21></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=75></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=75></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=61></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=4></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=10></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=75></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=75></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=62></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=1></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=12></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=5></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=59></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=1></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=11></TD>
		<TD><IMG height=1 src="Images/spacer.gif" width=16></TD>
	</TR>
	<TR>
		<TD><IMG height=20 src="Images/01_01.gif" width=108></TD>
		<TD colSpan=4><IMG height=20 src="Images/01_02.gif" width=105></TD>
		<TD colSpan=4><IMG height=20 src="Images/01_03.gif" width=232></TD>
		<TD colSpan=5><IMG height=20 src="Images/01_04.gif" width=226></TD>
		<TD colSpan=7><IMG height=20 src="Images/01_05.gif" width=105></TD>
	</TR>
  <TR>
		<TD rowSpan=4><IMG src="Images/01_06.gif"></TD>
		<TD colSpan=4 rowSpan=3><IMG src="Images/01_07.gif"></TD>
		<TD colSpan=10 rowSpan=3>
			<TABLE height=67 cellSpacing=0 cellPadding=0 width=459 border=0>
				<TR>
					<TD><IMG height=67 src="Images/logo.gif" width=459 border=0></TD>
				</TR>
			</TABLE>
		</TD>
		<TD colSpan=2 rowSpan=4><IMG height=68 src="Images/01_09.gif" width=17></TD>
		<TD align=center background="Images/top-1.gif" height=24 width=59 style="padding-top: 4px"><a href="QA.htm" target="_blank"><font class=A1>��������</font></a></TD><%'���� <A onmouseover="MM_showHideLayers('document.layers[\'Layer1\']','document.all[\'Layer1\']','show','document.layers[\'Layer2\']','document.all[\'Layer2\']','hide')" href="http://192.168.14.1/ftp/1/1/#"><IMG height=24 src="Images/01_10.gif" width=59 border=0></A></TD>%>
		<TD colSpan=3 rowSpan=2><IMG height=46 src="Images/01_11.gif" width=28></TD>
	</TR>
	<TR>
		<TD><IMG height=22 src="Images/01_12.gif" width=59></TD>
	</TR>
	<TR align=center>
		<%if Session.Contents("KEY")="super" then%>
			<TD colSpan=2 background="Images/top-2.gif" height=21 width=60 style="padding-top: 8px"><A onmouseover="MM_showHideLayers('document.layers[\'Layer1\']','document.all[\'Layer1\']','hide','document.layers[\'Layer2\']','document.all[\'Layer2\']','show')" href="http://192.168.14.1/ftp/1/1/#"><font class=A1>�������</font></A></TD>
		<%else%>
			<TD colSpan=2 background="Images/top-2.gif" height=21 width=60 style="padding-top: 8px"><font class=A1>�������</font></TD>
		<%end if%>
		<TD colSpan=2 rowSpan=2><IMG height=22 src="Images/01_14.gif" width=27></TD>
	</TR>
	<TR>
		<TD><IMG height=1 src="Images/01_15.gif" width=1></TD>
		<TD colSpan=3 rowSpan=2><IMG src="Images/01_16.gif"></TD>
		<TD colSpan=5 rowSpan=2><IMG height=19 src="Images/01_17.gif" width=236></TD>
		<TD colSpan=5 rowSpan=2><IMG height=19 src="Images/01_18.gif" width=223></TD>
		<TD colSpan=2><IMG height=1 src="Images/01_19.gif" width=60></TD>
	</TR>
	<TR>
		<TD colSpan=2><IMG src="Images/01_20.gif"></TD>
		<TD colSpan=6><IMG height=18 src="Images/01_21.gif" width=104></TD>
	</TR>
	<TR align=center>
		<TD background="Images/top_index.gif" height=33 width=159 style="padding-top: 6px" colSpan=3><a href=default.asp>&nbsp;&nbsp;<font class=A1>��ҳ</font></a><font class=A1>&nbsp;>>>>>&nbsp;</font><a href=Gstbook.asp><font class=A1>���Ա�</font></a></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp"><font class=A1>����ר��</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  			 ><A href="default.asp?classid=80"><font class=A1>��������</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px" 			><A href="default.asp?classid=1" ><font class=A1>���ĸ���</A></font></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp?classid=2" ><font class=A1>ŷ������</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  			 ><A href="default.asp?classid=3" ><font class=A1>�պ�����</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px" 			 ><A href="default.asp?classid=4" ><font class=A1>��������</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp?classid=0" ><font class=A1>���߰���</font></A></TD>
		<TD background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=4><A href="default.asp?classid=87"><font class=A1>ɧ���Ƽ�</font></A></TD>
		<TD><IMG height=33 src="Images/01_31.gif" width=16></TD>
	</TR>
	<TR>
		<TD colSpan=4><IMG height=15 src="Images/01_32.gif" width=195></TD>
		<TD colSpan=8><IMG height=15 src="Images/01_33.gif" width=339></TD>
		<TD colSpan=9><IMG height=15 src="Images/01_34.gif" width=242></TD>
	</TR>
</TABLE>
</CENTER></DIV>

<%
select case cmd
	case "AddNew"
		gstbook_addnew
	case "Display"
		gstbook_display
	case else
		gstbook_main
end select
%>

<TABLE cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<TR>
		<TD><IMG src="Images/copyright_01.gif"></TD>
		<TD><IMG src="Images/copyright_pengjohn.GIF"></TD>
		<TD><IMG src="Images/copyright_04.gif"></TD>
	</TR>
</TABLE>
<DIV></DIV>
</BODY>
</HTML>
<!--#include file="conn_music_close.asp"-->

<%'------------------------------------���Ա�'------------------------------------%>
<%function gstbook_main%>
<TABLE class=TableLine cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<TR>
		<TD background=Images/02_07.gif height="100%">
      	<DIV align=center><CENTER>
				<table width=740>
					<tr>
						<td>
						<%
						sql="select * from gstbook order by gstbooklasttime desc, gstbookTopID desc, gstbookorder"
						rs.open sql,conn,1,1
						if rs.eof and rs.bof then
							showtoolbar(Admin)
							response.write "<p align='center'>�� �� �� �� ��</p>"
						else
							totalPut=rs.recordcount '��¼����
							TotalPages = (totalPut-1) \ MaxPerPage_GstBook + 1
							if currentPage=1 then
								showtoolbar(Admin)
								showContent
							else
								if (currentPage-1)*MaxPerPage_GstBook<totalPut then
									rs.move  (currentPage-1)*MaxPerPage_GstBook
									dim bookmark
									bookmark=rs.bookmark '�ƶ�����ʼ��ʾ�ļ�¼λ��
									showtoolbar(Admin)
									showContent
								else
									currentPage=1
									showtoolbar(Admin)
									showContent
								end if
							end if
						end if
						rs.close
						set rs=nothing  
						%>
						</td>
					</tr>
				</table>
			</CENTER></DIV>
		</TD>
	</TR>
</TABLE>
<%end function%>

<%'------------------------------------�����б�'------------------------------------%>
<%Sub ShowContent%>
	<div align="center"><center>
	<table border="0" width="100%" cellspacing="1">
		<tr>
			<td>��</td>
			<td>��</td>
			<td>��</td>
			<td>��</td>
		</tr>
		<tr>
			<td width="6%" align="center"><b>No</b></td>
			<td width="60%" align="center"><b>����</b></td>
			<td width="12%" align="center"><b>������</b></td>
			<td width="22%" align="center"><b>����ʱ��</b></td>
		</tr>
		<%
		dim i
		i=0
		Do While Not rs.EOF or err
			if rs("gstbooktype")=1 then
				response.write "<tr bgcolor=#880000>"
			else
				response.write "<tr bgcolor=#000000>"
			end if
			response.write "<td align=center>" &rs("gstbook_ID") &"</td>"
			response.write "<td>"
			temp=1
			do while temp<rs("gstbooklayer")
				if temp=1 then
					response.write("<img border=0 src=images/gb_blank.gif>")
				else
					response.write("<img border=0 src=images/gb_blank.gif>")
				end if
				temp=temp+1    
			loop
			if temp=1 then    
				response.write("��")    
			else
				response.write("��")
			end if
			response.write "<a href=gstbook.asp?cmd=Display&gstbookID="& rs("gstbook_ID")&">"
			gstbooktitle = rs("gstbooktitle")
			if len(gstbooktitle)>35 then
				gstbooktitle = left(gstbooktitle,35) & "����"
			end if
			response.write gstbooktitle&"</a></td>"
			response.write "<td align=center><A href=javascript:openScript('user_show.asp?username="&rs("gstbookname")&"',400,220)>"& rs("gstbookname") &"</td>"
			response.write "<td align=center>"& rs("gstbookdatetime") &"</td>"
			response.write "</tr>"
			i=i+1
			if i>=MaxPerPage_GstBook then exit do 'ѭ��ʱ�����β�������˳��������¼�ﵽҳ�����ʾ����Ҳ�˳�
		rs.movenext
		loop
		%>
	</tr>
</table>
</div></center>
<%End Sub%>

<%Sub ShowToolBar(IsAdmin)
	if Session.Contents("mp3_username")<>"" then
		response.write "<img border=0 src=images/edit.gif><a href=gstbook.asp?cmd=AddNew>����������</a>"
	else
	   response.write "<a href=default.asp>�������ȵ�¼</a>"
	end if
	Response.Write "��"
	Response.Write "�� <b>"&totalput&"</b> ������"
	Response.Write "�� <b>"&currentPage& "<b> / <b>" &TotalPages& "</b >ҳ������"
	If currentPage <> 1 Then
		Response.Write "<A HREF=gstbook.asp>��ҳ</A>&nbsp;&nbsp;"
		Response.Write "<A HREF=gstbook.asp?Page=" & (currentPage-1) & ">��һҳ</A>&nbsp;&nbsp;"
	else
		Response.Write "��ҳ&nbsp;&nbsp;"
		Response.Write "��һҳ&nbsp;&nbsp;"				
	End If
	If currentPage <> TotalPages Then
		Response.Write "<A HREF=gstbook.asp?Page=" & (currentPage+1) & ">��һҳ</A>&nbsp;&nbsp;"
		Response.Write "<A HREF=gstbook.asp?Page=" & TotalPages & ">βҳ</A>"
	else
		Response.Write "��һҳ&nbsp;&nbsp;"
		Response.Write "βҳ"				
	End If
	Response.Write "��"
End Sub
%>

<%'------------------------------------��������'------------------------------------%>
<%function gstbook_addnew%>
<TABLE class=TableLine cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<TR>
		<TD background=Images/02_07.gif height="100%">
      	<DIV align=center><CENTER>
				<table width=740>
				<form method="POST" action="gstbook_newplan.asp">
				<input type="hidden" name="action" size="15" value="new">
					<tr>
						<td>��</td>
					</tr>
					<tr>
						<td>��������ӣ�</td>
					</tr>
					<tr>
						<td>���������⣺<input name="gstbooktitle" size="40" style="border: 1px inset rgb(255,255,255)">  *�����30��</td>
					</tr>
					<tr>
						<td>��ѡ�����ͣ�
							<input type=radio value=�������� name=mode checked>����&nbsp;&nbsp;			
							<input type=radio value=������ name=mode>���&nbsp;&nbsp;
							<input type=radio value=���Ƽ��� name=mode>�Ƽ�&nbsp;&nbsp;
							<input type=radio value=�۹���� name=mode>����&nbsp;&nbsp;
							<input type=radio value=��ԭ���� name=mode>ԭ��&nbsp;&nbsp;
						</td>
					</tr>
					<tr>
						<td>����������<img src="images/edit.gif" width="18" height="13"> ��������д���������</td>
					</tr>
					<tr>
						<td>����������<textarea name="gstbookcontent" cols="65" rows="16" style="border: 1px inset rgb(255,255,255)"></textarea>
						</td>
					</tr>
					<tr>
						<td align=center>
							<input type="Image" src="images/button-ok.gif" tppabs title="ȷ��" name="Submit" WIDTH="73" HEIGHT="19">
						</td>
					</tr>
				</form>
				</table>
			</CENTER></DIV>
		</TD>
	</TR>
</TABLE>
<%end function%>

<%'------------------------------------��������'------------------------------------%>
<%function gstbook_display%>
<%
gstbookID = request("gstbookID")

set rs=server.createobject("adodb.recordset")
rs.open ("SELECT * FROM gstbook where gstbook_id=" & gstbookID),conn,1,1
	gstbookorder		= rs("gstbookorder")
	gstbookTopID		= rs("gstbookTopID")
	gstbooklayer		= rs("gstbooklayer")
	gstbooktitle		= changechr_level1(rs("gstbooktitle"))
	gstbookname			= rs("gstbookname")
	gstbookdatetime	= rs("gstbookdatetime")
	gstbookcontent		= changechr(rs("gstbookcontent"))
rs.close
%>
<TABLE class=TableLine cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<TR>
		<TD background=Images/02_07.gif height="100%">
      	<DIV align=center><CENTER>
			<table width=740>
				<tr>
					<td>����������<%=gstbookdatetime%>&nbsp;<a href="javascript:openScript('user_show.asp?username=<%=gstbookname%>',400,220)"><b><%=gstbookname%></b></a>&nbsp;����д����	</td>
				<tr>
					<td>���⣺<%=gstbooktitle%>	</td>
				</tr>
				<tr>
					<td>���ݣ�</td>
				</tr>
				<tr>
					<td>
						<table cellPadding=10>
							<tr>
								<td width=40></td>
								<td width=600 style="border: 1px outset rgb(255,255,255)"><%=gstbookcontent%>��</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
					<%if Session.Contents("IsAdmin")=true or Session.Contents("KEY")="check" or Session.Contents("mp3_username")=gstbookname then%>
						����	<a href="gstbook_del.asp?gstbookid=<%=gstbookID%>">��ɾ����</a>
					<%end if%>
					<%if Session.Contents("IsAdmin")=true or Session.Contents("KEY")="check" then%>			
						<a href="gstbook_best.asp?gstbookid=<%=gstbookID%>&mode=OK">����Ϊ������</a>
						<a href="gstbook_best.asp?gstbookid=<%=gstbookID%>&mode=NG">��ȡ��������</a>
					<%end if%>
					</td>
				</tr>
				<tr>
					<td height=20><hr size="1" color="#ffffff"></td>
				</tr>
				<%
				rs.open ("SELECT * FROM gstbook WHERE gstbookTopID=" & gstbookTopID & " ORDER BY gstbookorder"),conn,1,1 
				if rs.recordcount>=2 then
				%> 
				<tr>
					<td>
						<table width="700" border="0" cellspacing="0" cellpadding="0" align="center"> 
			  				<tr> 
								<td> 
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center"> 
										<tr>
											<td width="6%" align="center">No</td>  
											<td width="58%"><b>&nbsp;&nbsp;&nbsp;�������</b></td> 
											<td width="12%" align="center">������</td> 
											<td width="24%" align="center">����ʱ��</td> 
										</tr> 
										<%do while not rs.eof %>  
										<tr bgcolor="#000000">  
											<td align="center"><%=rs("gstbook_ID")%></td>
											<td>
											<%temp=1    
											do while temp<rs("gstbooklayer")
												if temp=1 then
													response.write("<img border=0 src=images/gb_blank.gif>")
												else
													response.write("<img border=0 src=images/gb_blank.gif>")
												end if
												temp=temp+1    
											loop
											if temp=1 then    
												response.write("��")
											else
												response.write("��")
											end if%>
											<a href="gstbook.asp?cmd=Display&gstbookID=<%=rs("gstbook_id")%>"><%=left(rs("gstbooktitle"),40)%></a>
											</td>
											<td align="center"><a href="javascript:openScript('user_show.asp?username=<%=rs("gstbookname")%>',400,220)"><%=left(rs("gstbookName"),10)%></a></td> 
											<td align="center"><%=rs("gstbookdatetime")%></td> 
										</tr> 
										<%rs.movenext 
										loop
										%>
									</table> 
								</td> 
							</tr> 
						</table>
					</td>
				</tr>
				<%end if
				rs.close 
				%>
				<tr>
					<td height=20><hr size="1" color="#ffffff"></td>
				</tr>
				<%if Session.Contents("mp3_username")<>"" then %> 
				<form action="gstbook_newplan.asp?action=follow&amp;gstbooklayer=<%=(gstbooklayer+1)%>&amp;gstbookTopID=<%=gstbookTopID%>&amp;gstbookorder=<%=(gstbookorder+1)%>" method="post"> 
					<tr>
						<td>
							��Ӧ���⣺<input name="gstbooktitle" size="50" value="Re:<%=gstbooktitle%>" style="border: 1px inset rgb(255,255,255)">  *�����30��
						</td>
					</tr> 
					<tr>  
						<td>����������<img src="images/edit.gif" width="18" height="13">��������д�����Ļظ�</td> 
					</tr> 
					<tr>  
						<td>����������<textarea name="gstbookcontent" cols="65" rows="12" style="border: 1px inset rgb(255,255,255)"></textarea></td> 
					</tr> 
					<tr>  
						<td height="40" align="center">
							<input type="Image" src="images/button-ok.gif" tppabs title="ȷ��" name="Submit" WIDTH="73" HEIGHT="19">
						</td>
					</tr> 
				</form> 
				<%end if%> 
			</TABLE>
			</CENTER></DIV>
		</TD>
	</TR>
</TABLE>


<%end function%>