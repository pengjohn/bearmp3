<!--#include file="ip_check.asp"-->
<!--#include file="conn_music_open.asp"-->
<!--#include file="conn_define.asp"-->
<!--#include file="conn_function.asp"-->

<%
Session.Timeout=600
dim totalPut   
dim CurrentPage
dim k
dim keyword
dim sql,rs
dim classname,classid
dim singername,singerid
dim albumid,albumname
dim titlename
dim cenbar

classid	= request("classid")
singerid	= request("singerid")
albumid	= request("albumid")
keyword	= request("keyword")			'��������
goodcomment_id = request("goodcomment_id")

if not isempty(request("page")) then
	currentPage = cint(request("page"))
else
	currentPage = 1
end if

set rs = server.createobject("adodb.recordset")

'����
if classid<>"" then
	rs.open "select * from class where classid="&cstr(classid),conn,1,1
	if not rs.eof then
		classname = rs("classname")
	end if
	rs.close
end if

'����
if singerid<>"" then
	rs.open "select * from singer,class where singer.classid = class.classid and singerid="&cstr(singerid),conn,1,1
	if not rs.eof then
		singername = rs("singername")
		classname  = rs("classname")
	end if
	rs.close
end if

'ר��
if albumid<>"" then
	rs.open "select * from album,singer,class where album.singerid=singer.singerid and singer.classid=class.classid and albumid="&cstr(albumid),conn,1,1
	if not rs.eof then
		albumname	= rs("albumname")
		singername	= rs("singername")
		classname	= rs("classname")
		albumcover	= rs("albumcover")
		albumname	= rs("albumname")
	end if
	rs.close
end if

if Session.Contents("mp3_username")<>"" then
	username = Session.Contents("mp3_username")
	sql="select * from user where username='"&username&"'" 
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		userscore = 0			'�޴�user
	else
		userscore = rs("userscore")	'��user����
		userface  = rs("userimage")
	end if
	rs.close
end if
%>


<html>
<head>
<title><%=Info_Title%><%=Info_Version%></title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">

<script language=JavaScript>
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

</head>

<body bgColor=#000000 leftMargin=0 topMargin=0 align="center">
<div align=center><center>
<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width=776 border=0>
	<tr>
		<td><img height=1 src="Images/spacer.gif" width=108></td>
		<td><img height=1 src="Images/spacer.gif" width=1></td>
		<td><img height=1 src="Images/spacer.gif" width=50></td>
		<td><img height=1 src="Images/spacer.gif" width=36></td>
		<td><img height=1 src="Images/spacer.gif" width=18></td>
		<td><img height=1 src="Images/spacer.gif" width=21></td>
		<td><img height=1 src="Images/spacer.gif" width=75></td>
		<td><img height=1 src="Images/spacer.gif" width=75></td>
		<td><img height=1 src="Images/spacer.gif" width=61></td>
		<td><img height=1 src="Images/spacer.gif" width=4></td>
		<td><img height=1 src="Images/spacer.gif" width=10></td>
		<td><img height=1 src="Images/spacer.gif" width=75></td>
		<td><img height=1 src="Images/spacer.gif" width=75></td>
		<td><img height=1 src="Images/spacer.gif" width=62></td>
		<td><img height=1 src="Images/spacer.gif" width=1></td>
		<td><img height=1 src="Images/spacer.gif" width=12></td>
		<td><img height=1 src="Images/spacer.gif" width=5></td>
		<td><img height=1 src="Images/spacer.gif" width=59></td>
		<td><img height=1 src="Images/spacer.gif" width=1></td>
		<td><img height=1 src="Images/spacer.gif" width=11></td>
		<td><img height=1 src="Images/spacer.gif" width=16></td>
	</tr>
	<tr>
		<td><img height=20 src="Images/01_01.gif" width=108></td>
		<td colSpan=4><img height=20 src="Images/01_02.gif" width=105></td>
		<td colSpan=4><img height=20 src="Images/01_03.gif" width=232></td>
		<td colSpan=5><img height=20 src="Images/01_04.gif" width=226></td>
		<td colSpan=7><img height=20 src="Images/01_05.gif" width=105></td>
	</tr>
  <tr>
		<td rowSpan=4><img src="Images/01_06.gif"></td>
		<td colSpan=4 rowSpan=3><img src="Images/01_07.gif"></td>
		<td colSpan=10 rowSpan=3>
			<table height=67 cellSpacing=0 cellPadding=0 width=459 border=0>
				<tr>
					<td><img height=67 src="Images/logo.gif" width=459 border=0></td>
				</tr>
			</table>
		</td>
		<td colSpan=2 rowSpan=4><img height=68 src="Images/01_09.gif" width=17></td>
		<td align=center background="Images/top-1.gif" height=24 width=59 style="padding-top: 4px"><a href="QA.htm" target="_blank"><font class=A1>��������</font></a></td>
		<td colSpan=3 rowSpan=2><img height=46 src="Images/01_11.gif" width=28></td>
	</tr>
	<tr>
		<td><img height=22 src="Images/01_12.gif" width=59></td>
	</tr>
	<tr align=center>
		<td colSpan=2 background="Images/top-2.gif" height=21 width=60 style="padding-top: 7px"><font class=A1>Ԥ����Ŀ</font></td>
		<td colSpan=2 rowSpan=2><img height=22 src="Images/01_14.gif" width=27></td>
	</tr>
	<tr>
		<td><img height=1 src="Images/01_15.gif" width=1></td>
		<td colSpan=3 rowSpan=2><img src="Images/01_16.gif"></td>
		<td colSpan=5 rowSpan=2><img height=19 src="Images/01_17.gif" width=236></td>
		<td colSpan=5 rowSpan=2><img height=19 src="Images/01_18.gif" width=223></td>
		<td colSpan=2><img height=1 src="Images/01_19.gif" width=60></td>
	</tr>
	<tr>
		<td colSpan=2><img src="Images/01_20.gif"></td>
		<td colSpan=6><img height=18 src="Images/01_21.gif" width=104></td>
	</tr>
	<tr align=center>
		<td background="Images/top_index.gif" height=33 width=159 style="padding-top: 6px" colSpan=3>&nbsp;&nbsp;<a href=default.asp><font class=A1>��ҳ</font></a><font class=A1>&nbsp;>>>>>&nbsp;</font><a href=Gstbook.asp><font class=A1>���Ա�</font></a></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp"><font class=A1>����ר��</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  			 ><A href="default.asp?classid=80"><font class=A1>��������</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px" 			><A href="default.asp?classid=1" ><font class=A1>���ĸ���</A></font></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp?classid=2" ><font class=A1>ŷ������</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  			 ><A href="default.asp?classid=3" ><font class=A1>�պ�����</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px" 			 ><A href="default.asp?classid=4" ><font class=A1>��������</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=3><A href="default.asp?classid=0" ><font class=A1>���߰���</font></A></td>
		<td background="Images/top_blank.gif" height=33 width=76 style="padding-top: 6px"  colSpan=4><A href="default.asp?classid=87"><font class=A1>�����Ƽ�</font></A></td>
		<td><img height=33 src="Images/01_31.gif" width=16></td>
	</tr>
	<tr>
		<td colSpan=4><img height=15 src="Images/01_32.gif" width=195></td>
		<td colSpan=8><img height=15 src="Images/01_33.gif" width=339></td>
		<td colSpan=9><img height=15 src="Images/01_34.gif" width=242></td>
	</tr>
</table>
</center></div>

<table class=TableLine cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<tr>
		<td vAlign=top width=185 background=Images/02_05.gif height="100%">
			<table style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=185 border=0>
				<tr>
					<td height="33" align=center vAlign=bottom background=Images/02_title.gif><font class=A1>��վ����</font></td>
				</tr>
				<tr>
					<td width=184 background=Images/02_02.gif align=center>
						<table style="BORDER-COLLAPSE: collapse" borderColor=#111111 height=1 cellSpacing=0 cellPadding=0 width="82%" border=0>
							<tr>
							<%
   						sql="select * from history order by historydate desc"
   						rs.open sql,conn,1,1
   						if not(rs.eof) then
   							history = rs("historycontent")
   						else
   							history = ""
   						end if
   						rs.close
   						%>
								<td class=jnfont4 width="51%" height=1><%=history%></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td width=185><img src="Images/02_04.gif" border=0><br><br></td>
				</tr>
			</table>
			<%if Session.Contents("KEY")="super" then%>
			<table style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=185 border=0>
				<tr>
					<td height="33" align=center vAlign=bottom background=Images/02_title.gif><font class=A1>�������</font></td>
				</tr>
				<tr>
					<td width=184 background=Images/02_02.gif align=center>
						<table style="BORDER-COLLAPSE: collapse" borderColor=#111111 height=1 cellSpacing=0 cellPadding=0 width="82%" border=0>
							<tr>
								<a href="album_add.asp?step=CLASS"><li>���ר��</li></a>
								<a href="song_add.asp?step=CLASS"><li>��ӵ���</li></a>
								<a href="admin_link.asp"><li>��ҳ���ӹ���</li></a>
								<a href="admin_history.asp"><li>��ҳ�������</li></a>
								<a href="admin_software.asp"><li>�������ع���</li></a>
								<a href="admin_goodcomment.asp"><li>��ҳ�Ƽ�����</li></a>

								<a href=admin_list.asp?list_mode=NEW_ALBUM><li>����100��ר��</li></a>
								<a href=admin_list.asp?list_mode=MISS_SONG><li>��ȱ����</li></a>
								<a href=admin_list.asp?list_mode=MISS_COVER><li>��ȱר������</li></a>
								<a href=admin_list.asp?list_mode=ALL_ALBUM><li>����ר��</li></a>
								<a href=admin_list.asp?list_mode=ALL_SONG><li>���и���</li></a>
								<a href=admin_list.asp?list_mode=ALL_SONG_LISTPRO><li>���и���ListPro</li></a>
								
								<a href="admin_check_song_link.asp"><li>�Զ�����������</li><br></a>
								<a href="admin_user_login_list.asp" target="_blank"><li>������Ϣ</li></a>
								<a href="admin_list_user.asp" target="_blank"><li>�û���½��Ϣ</li></a>
								<a href="ip_admin.asp" target="_blank"><li>IPȨ�޹���</li></a>
								<a href="gstbook_delall.asp"><li>�����������</li></a>
								<a href="admin_manage.asp"><li>վ�����</li></a>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td width=185><img src="Images/02_04.gif" border=0><BR>��</td>
				</tr>
			</table>
			<%end if%>
			<table cellSpacing=0 cellPadding=0 width=185 background=Images/02_05.gif border=0>
				<tr>
					<td width="100%" align=center>
						<table cellSpacing=0 cellPadding=0 width="100%" border=0>
							<tr>
								<td align=middle width="100%"><img src="Images/bar_user.gif"></td>
							</tr>
						</table>
						<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=3 cellPadding=0 width=160 bgColor=#055192 border=0>
						<%if Session.Contents("mp3_username")<>"" then%>
							<tr>
								<td>
									<img src="<%=userface%>" align="left" width="64" height="64">
									<%=Session.Contents("mp3_username")%><br>
									��Ļ��� [ <b><%=userscore%></b> ]<br>
									<a href="user_edit.asp">�޸ĸ�������</a>
								</td>
							</tr>
						<%else%>
							<FORM method="post" action="user_chklogin.asp">
							<tr>
								<td class=jnfont4 align=right width="33%">�� �ţ�</td>
								<td><INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #5a9ef1 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; PADDING-TOP: 1px; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" size=12 name="username"></td>
							</tr>
							<tr>
								<td class=jnfont4 align=right width="33%">�� �룺</td>
								<td><INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #5a9ef1 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; PADDING-TOP: 1px; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" type=password size=12 name="password"></td>
							</tr>
							<tr>
								<td align=middle width="100%" colSpan=2 height=35>
									<INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; BORDER-TOP: #5a9ef1 1px solid; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" type=submit value=��½ name=submit>
									<INPUT style="BORDER-RIGHT: #1e9dfb 1px solid; BORDER-TOP: #1e9dfb 1px solid; BORDER-LEFT: #1e9dfb 1px solid; COLOR: #c5e6fe; BORDER-BOTTOM: #1e9dfb 1px solid; BACKGROUND-COLOR: #055192" onclick="javascript:window.open('user_reg.asp','_self','')" type=button value=ע�� name=reg>
								</td>
							</tr>
							</FORM>
						<%end if%>
						</table>
					</td>
				</tr>
				<tr>
					<td width="100%" align=center>
						<table cellSpacing=0 cellPadding=0 width="100%" border=0>
							<tr>
								<td width="100%" align=center><img src="Images/barbar.gif" border=0><BR><img height=22 src="Images/bar3.gif" width=160 border=0></td>
							</tr>
						</table>
						<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=160 bgColor=#055192 border=0>
						<form name="form1" method="post" action="default.asp">
							<tr>
								<td class=jnfont4 align=right width="45%">��&nbsp; ����</td>
								<td>
									<SELECT style="FONT-SIZE: 9pt;COLOR: #c5e6fe; BACKGROUND-COLOR: #055192" size=1 name=classid>
									<OPTION value="81" selected>==�� ��==</OPTION>
									<OPTION value="82">==�� ��==</OPTION>
									<OPTION value="83">==ר ��==</OPTION>
									<OPTION value="84">==�� ��==</OPTION>
									</SELECT>
								</td>
							</tr>
							<tr>
								<td class=jnfont4 align=right width="45%">�ؼ��֣�</td>
								<td><INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: #5a9ef1 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; PADDING-TOP: 1px; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" size=10 name=keyword></td>
							</tr>
							<tr>
								<td align=middle width="100%" colSpan=2 height=30>
									<INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; BORDER-TOP: #5a9ef1 1px solid; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" type=submit value="�� ��" name=submit> 
									<INPUT style="BORDER-RIGHT: #5a9ef1 1px solid; BORDER-TOP: #5a9ef1 1px solid; BORDER-LEFT: #5a9ef1 1px solid; COLOR: #c5e6fe; BORDER-BOTTOM: #5a9ef1 1px solid; BACKGROUND-COLOR: #055192" type=reset value="�� ��" name=submit1>
								</td>
							</tr>
						</FORM>
						</table>
					</td>
				</tr>
				<tr>
					<td width="100%" align=center><img src="Images/barbar.gif" border=0></td>
				</tr>
<%if Switch_Tools = 1 then%>
				<tr>
					<td align=center>
						<table cellSpacing=0 cellPadding=0 width=160 border=0>
							<tr>
								<td width="100%"><img src="Images/bar5.gif" border=0></td>
							</tr>
							<tr>
								<td width="100%">
									<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" borderColor=#111111 height=1 cellSpacing=0 cellPadding=4 width="100%" bgColor=#055192 border=0>
										<%
										sql="select * from software order by software_id desc;"  
										rs.Open sql,conn,1,1
										do while not(rs.eof) 
										%>
										<tr>
											<td align=left width="100%" height=40>
												<li><a href="<%=rs("software_dl")%>"><%=rs("software_title")%></a></li><br>
												<%=rs("software_memo")%>
											</td>
										</tr>
										<%
										rs.movenext
										loop
										rs.close
										%>
									</table>
								</td>
							</tr>
							<tr>
								<td width="100%"><img src="Images/barbar.gif" border=0></td>
							</tr>
						</table>
					</td>
				</tr>
<%end if%>				
				
			</table>
		</td>
		
		<td vAlign=top align=middle width=405>
			<%'�����Ƽ�---------------------------------%>
			<%if classid="" and singerid="" and albumid="" then
				goodcomment_show
			end if
			%>

			<%'�б����壺���֡�ר����������------------%>
			<%main_show%>

			<%'���Ա�---------------------------------%>
			<%if Switch_GuestBook=1 then
				guestbook_show
			end if
			%>
		</td>

	<td vAlign=top width=188 background=Images/04_02.gif>
		<table cellSpacing=0 cellPadding=0 width=188 border=0>
		<%'��������---------------------------------%>
		<%if switch_TopSong=1 then
			TopSongs_Show
		end if
		%>
		
		<%'ר������---------------------------------%>
		<%if Switch_TopAlbum=1 then
			TopAlbum_Show
		end if
		%>
		
		<%'ͳ������---------------------------------%>
		<%if Switch_Count = 1 then
		   TotalCount_Show
		end if
		%>

		<% if Switch_Link = 1 then%>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar4.gif" border=0></td>
			</tr>
			<tr>
				<td width="100%">
					<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" cellPadding=3 width=166 align=center bgColor=#055192 border=0>
						<tr>
							<td vAlign=top width="100%">
									<%
									k=0
									sql="select * from link order by link_id desc;"  
									rs.Open sql,conn,1,1
									do while not(rs.eof) 
										response.write "<li><a href='"&rs("link_address")&"'> "&rs("link_title")&"</a><br>"
									rs.movenext
									loop
									rs.close
									%>
							</td>
						</tr>
						<tr>
							<td width="100%"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar1-2.gif" border=0></td>
			</tr>
			<%end if%>
		</table>
	</td>
</tr>
</table>
<table cellSpacing=0 cellPadding=0 width=776 align=center border=0>
	<tr>
		<td><img src="Images/copyright_01.gif"></td>
		<td><img src="Images/copyright_pengjohn.GIF"></td>
		<td><img src="Images/copyright_04.gif"></td>
	</tr>
</table>
</body>
</html>
<%set rs = nothing%>
<!--#include file="conn_music_close.asp"-->

<%'------------------------------------�б�����'------------------------------------%>
<%sub main_show%>
<table cellSpacing=0 cellPadding=0 width=403 background=Images/cenbar2.gif border=0>
	<tr>
		<td>
<% '��������classid��00-21, 80-99Ϊ������չ 
	select case classid
		case ""
			if singerid<>"" then
				sql="select * from album where singerid="+cstr(singerid)+" order by ((albumid mod 100)<>0),albumname" 
				cenbar = "Images/cenbar_listalbum.GIF"
				listalbum
			else		'singer=0
				if albumid<>"" then
					sql="select * from song where albumid="+cstr(albumid)+" order by songid" 
					cenbar = "Images/cenbar_listalbumsongs.GIF"		'ר����Ŀ�б�
					listalbumsongs
				else		'albumid=0
					sql="select * from album where (albumid mod 100<>0) order by album_id desc"
					titlename = "����ר��"
					cenbar = "Images/cenbar_newablum.gif"				'����ר��"
					newalbum_list
				end if
			end if
		case 80
			sql="select * from song,album where song.albumid=album.albumid order by song_id desc"
			titlename = "��������"
			cenbar = "Images/cenbar_newsong.gif"				'��������"
			listsongs
		case 81
			sql="select * from song,album where song.albumid=album.albumid and song.songname Like '%"& keyword &"%' order by song.song_id desc"
			titlename = "�����ؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'��Ŀ�ؼ���
			listsongs
		case 82
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername Like '%"& keyword &"%' order by album.albumid desc"
			titlename = "���ֹؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'���ֹؼ���
			listsongs
		case 83
			sql="select * from song,album where song.albumid=album.albumid and album.albumname Like '%"& keyword &"%' order by album.singerid,album.albumid"
			titlename = "ר���ؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'ר���ؼ���
			listsongs
		case 84
			sql="select * from song,album where song.albumid=album.albumid and songlrc Like '%"& keyword &"%' order by song_id desc"
			titlename = "��ʹؼ��֡�<font color=#00FF00>"&keyword&"</font>������"
			cenbar = "Images/cenbar_result.GIF"	'��ʹؼ���
			listsongs
		case 86
			sql="select * from song,album where song.albumid=album.albumid order by songhit desc"
			titlename = "��������"
			listsongs
		case 87
			titlename = "�����Ƽ�"
			goodcomment_list
		case 88
			sql="select * from song,album where song.albumid=album.albumid and songfile='' order by song.albumid desc"
			titlename = "���͸���"
			listsongs
		case 89
			sql="select * from album where albumcover='images/nocover.gif' order by albumname" 
			titlename = "����CDר������"
			listalbum
		case 90
			sql="select * from song,album,singer where album.albumid=song.albumid and song.singerid=singer.singerid and singer.singername='"& keyword &"' order by album.albumid desc"
			titlename = "���֡�<font color=#00FF00>"&keyword&"</font>�����и����б�"
			cenbar = "Images/cenbar_listalbumsongs.GIF"	'�������и����б�
			listsongs
		case 91
			sql="select * from song,album where song.albumid=album.albumid and songhot>=5 order by song.albumid desc"
			titlename = "�Ƽ�����"
			cenbar = "Images/cenbar_best.GIF"	'�Ƽ�����
			listsongs
		case 100
			sql="select * from user order by userid desc"
			user_list
		case 0
			cenbar = "Images/cenbar_else.gif"				'��������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 1
			cenbar = "Images/cenbar_cn.gif"					'�������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 2
			cenbar = "Images/cenbar_en.gif"					'ŷ������
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 3
			cenbar = "Images/cenbar_ja_co.gif"				'�պ�����
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case 4
			cenbar = "Images/cenbar_NewAge.gif"				'New Age
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
		case else
			sql="select * from singer where classid="+cstr(classid)+" order by singerindex, singername" 
			listsinger
	end select					
	%>
		</td>
	</tr>
	<tr>
		<td vAlign=top width="100%"><img src="Images/cenbar3.gif" border=0></td>
	</tr>
</table>
<%end sub%>

<%'-----------------------------------���Ա��б�-----------------------------------%>
<%sub guestbook_show%>
<table cellSpacing=0 cellPadding=0 width=403 background=Images/cenbar2.gif border=0>
	<tr>
		<td>
			<table cellSpacing=0 cellPadding=0 width=403 border=0>
				<tr>
					<td vAlign=top width=403 background="images/cenbar_guestbook.gif" height=33>&nbsp;</td>
				</tr>
				<tr>
					<td align=center>
						<table height=81 cellSpacing=0 cellPadding=2 width="97%" border=0>
							<%
							k=0
							sql="select * from gstbook order by gstbook_id desc;"  
							rs.Open sql,conn,1,1
							do while not(rs.eof) 
							%>
							<tr>
								<td width="38%" align=right><a href="javascript:openScript('user_show.asp?username=<%=rs("gstbookname")%>',400,220)"><b><%=rs("gstbookname")%></b></a><br>
								 [<%=rs("gstbookdatetime")%>]
								</td>
								<td width="4%" vAlign=top><img src="images/gbook.gif"></td>
								<td width="58%" vAlign=top>
									<%
									gstbooktitle = rs("gstbooktitle")
									if len(gstbooktitle)>20 then
										gstbooktitle = left(gstbooktitle,20) & "����"
									end if
									%>
									<a href="gstbook.asp?cmd=Display&gstbookID=<%=rs("gstbook_ID")%>"><%=gstbooktitle%></a>
								</td>
							</tr>
							<%
							k=k+1
							if k>=Num_Gstbook then exit do
							rs.movenext
							loop
							rs.close
							%>
							<tr>
								<td width="100%" align="right" colspan="3"><a href="gstbook.asp">�� �������Ա�</a></td>
							</tr>   
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td vAlign=top width="100%"><img src="Images/cenbar3.gif" border=0></td>
	</tr>
</table>
<%end sub%>

<%'------------------------------------�����б�'------------------------------------%>
<%sub listsinger%>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
	<%
	dim i 
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		response.write "<p align='center'>��ʱû���ռ�</p>" 
	end if
	%>
				<tr>
					<td width="100%" colspan="<%=MaxPerRow_Singer%>">��ǰλ�ã�<%=classname%> �� �����б�</td>
				</tr>
				<tr>
					<td width="100%" height="20" colspan="<%=MaxPerRow_Singer%>"></td>
				</tr>
				<tr>
		<%
		i=1
		preindex = ""
		do while not rs.eof
		if rs("singerindex")<>preindex then
			if i<MaxPerRow_Singer then
				colspannum = MaxPerRow_Singer-i+1
				response.write "<td colspan='"& colspannum &"' width='"& (colspannum*100/MaxPerRow_Singer) &"%'></td>"
			end if
			response.write "</tr><tr><td colspan='"&MaxPerRow_Singer&"' height='1' bgcolor='#FFFFFF'></td></tr>"
			response.write "<tr><td colspan='"&MaxPerRow_Singer&"'><b>��"&rs("singerindex")&"��</b></td></tr><tr>"
			preindex = rs("singerindex")
			i=1
		end if
		response.write "<td width="&(100/MaxPerRow_Singer)&"% align=left>��<a href='default.asp?singerid="&rs("singerid")&"'>"&rs("singername")&"</a></td>"
		if i>=MaxPerRow_Singer then
			i=1
			response.write "</tr><tr>"
		else
			i=i+1
		end if
	  	rs.movenext
		loop
		rs.close
		response.write "</tr><tr><td colspan='"&MaxPerRow_Singer&"'><hr size=1 color=#ffffff></td></tr>"
		%>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end sub%>

<%''------------------------------------ר���б�'------------------------------------%>
<%sub listalbum%>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
			<%
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then 
				response.write "<p align='center'>��ʱû���ռ�</p>" 
			else 
			%>
				<tr>
					<td>��ǰλ�ã�<a href="default.asp?classid=<%=(singerid\10^Mul_Singer)%>"><%=classname%></a> �� ���֣�<%=singername%>��ר���б�����<font color="red"><b><%=rs.recordcount%></b></font>����</td>
				</tr>
				<tr>
					<td>
					<%if Session.Contents("IsAdmin")=true then%>
						���������
						<a href="singer_edit.asp?singerid=<%=singerid%>" target=_blank>�۱༭�������ϣ�</a>
						<a href="album_add.asp?step=ALBUM&singerid=<%=singerid%>">�����ר����</a>
					<%end if%>
					</td>
				</tr>
				<tr>
					<td><a href="default.asp?classid=90&keyword=<%=singername%>">���и����б�</a></td>
				</tr>
				<tr>
					<td>ר���б�</td>
				</tr>
				<tr>
					<td colspan="2">
						<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
							<tr align="center">
								<td width="12%" height="22" background="images/dh-di.jpg">ר�����</td>
								<td width="40%" background="images/dh-di.jpg">CDר��</td>
								<td width="15%" background="images/dh-di.jpg">��������</td>
								<%if Switch_Comment=1 then
									response.write "<td width='12%' background='images/dh-di.jpg'>ר������</td>"
									response.write "<td width='15%' background='images/dh-di.jpg'>����</td>"
								end if
								%>
							</tr>
							<%do while not rs.eof%>
							<tr align="center">
								<td><%=rs("albumid")%></a></td>
								<td align="left"><a href="default.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a></td>
								<td><%=rs("albumdatetime")%></td>
								<%if Switch_Comment=1 then%>
									<td><a href="comment_show.asp?commentindex=<%=rs("albumid")%>">�鿴</a></td>
									<td>
										<%if rs("albumgrade")=0 then
											response.write "--"
										elseif rs("albumgrade")=1 then
											response.write "<img src='images/bank1.gif'>"
										elseif rs("albumgrade")=2 then
											response.write "<img src='images/bank2.gif'>"
										elseif rs("albumgrade")=3 then
											response.write "<img src='images/bank3.gif'>"
										elseif rs("albumgrade")=4 then
											response.write "<img src='images/bank4.gif'>"
										elseif rs("albumgrade")=5 then
											response.write "<img src='images/bank5.gif'>"
										end if
										%>
									</td>
								<%end if%>
							</tr>
							<%
							rs.movenext
							loop
							%>
						</table>
					</td>
				</tr>
				<%end if
				rs.close
				%>
			</table>
		</td>
	</tr>
</table>
<%end sub%>

<%''------------------------------------ר�������б�'------------------------------------%>
<% sub listalbumsongs %>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
				<tr>
					<td>��ǰλ�ã�<a href="default.asp?classid=<%=(albumid\10^(Mul_Singer+Mul_Album))%>"><%=classname%></a> �� <a href="default.asp?singerid=<%=(albumid\10^Mul_Album)%>"><%=singername%></a> �� ר����<%=albumname%>�ݸ����б�</td>
				</tr>
				<%if Session.Contents("IsAdmin")=true then%>
				<tr>
					<td>���������
					<a href=song_add.asp?step=SONG&albumid=<%=albumid%>>�������Ŀ��</a>
					<a href=album_edit.asp?albumid=<%=albumid%>>���޸ģ�</a>
					<a href=album_del.asp?albumid=<%=albumid%>>��ɾ����</a>
					<a href=album_refresh.asp?albumid=<%=albumid%>>���Զ����£�</a>
					<br><br>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td width="100%" align="center"><br><br>
						<%show_cd(albumcover)%>
						<br><br><font color="white"><%=albumname%></font><br><br>
					</td>
				</tr>
				<tr>
					<td>
						<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
							<tr align="center">
								<td width="25%" height="22" background="images/dh-di.jpg">�������</td>
								<td width="45%" background="images/dh-di.jpg">��������</td>
								<%if Switch_Comment=1 then
									response.write "<td width=10% background=images/dh-di.jpg>����</td>"
								end if%>
								<td width="15%" background="images/dh-di.jpg">�������</td>
								<td width="5%"  background="images/dh-di.jpg">��</td>
								<td width="5%"  background="images/dh-di.jpg">��</td>
								<td width="5%"  background="images/dh-di.jpg">��</td>
								<%if Session.Contents("IsAdmin")=true then%>
									<td width="5%" background="images/dh-di.jpg">ɾ</td>
									<td width="5%" background="images/dh-di.jpg">��</td>
									<td width="5%" background="images/dh-di.jpg">��</td>
								<%end if%>
							</tr>
							<%
							rs.open sql,conn,1,1
							%>
							����<font color="red"><b><%=rs.recordcount%></b></font>����
							<%do while not rs.eof
							%>
							<tr align="center">
								<td><%=rs("songid")%></td>
								<td>
									<%if rs("songhot")>=5 then
										response.write "<font color=red>��</font>"
									end if
									if rs("songfile")<>"" then
										response.write rs("songname")
									else
										response.write rs("songname")&"(ȱ)"
									end if%>
								</td>
								<%if Switch_Comment=1 then
									response.write "<td><a href='comment_show.asp?commentindex="&rs("songid")&"'>�鿴</a></td>"
								end if%>
								<td><%=rs("songdatetime")%></td>
								<td><a href="lrc_show.asp?song_id=<%=rs("song_id")%>" target="_blank">
									<%if rs("songlrc")<>"" then
										response.write "<img src=images/lrc.gif border=0 alt=���>"
									else
										response.write "��"
									end if%>
								</td>
								<%if rs("songfile")<>"" then%>
									<%if Mode_Download=1 then%>
										<td><a href=song_download.asp?song_id=<%=rs("song_id")%>><img src="images/download.gif" border="0" alt="����"></a></td>
									<%else%>
										<td><a href="<%=rs("songfile")%>"><img src="images/download.gif" border="0" alt="����"></a></td>
									<%end if%>
									<%select case Mode_Play
										case 1
									%>
										<td><a href="javascript:playonline('play01.asp?song_id=<%=rs("song_id")%>',388,145)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>
									<% case 2%>
										<td><a href="javascript:playonline('play02.asp?song_id=<%=rs("song_id")%>',380,400)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>
									<% case else%>
										<td><a href="javascript:playonline('play01.asp?song_id=<%=rs("song_id")%>',388,145)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>									
									<%end select%>
								<%else%>
									<td>��</td>
									<td>��</td>
								<%end if%>
								<%if Session.Contents("IsAdmin")=true then%>
									<td><a href="song_del.asp?song_id=<%=rs("song_id")%>"><img src="images/del.gif" border="0" alt="ɾ"></a></td>
									<td><a href="song_best.asp?song_id=<%=rs("song_id")%>"><img src="images/best.gif" border="0" alt="�Ƽ�"></a></td>
									<td><a href="song_edit.asp?song_id=<%=rs("song_id")%>" target="_blank"><img src="images/edit.gif" border="0" alt="�༭"></a></td>
								<%end if%>
							</tr>
							<%
							rs.movenext
							loop
							rs.close
							%>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% end sub %>

<%'------------------------------------ȫ�����б�------------------------------------%>
<%sub listsongs%>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>��ʱû���ռ�</p>"
	else
		totalPut=rs.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage_Songs>totalput then
			if (totalPut mod MaxPerPage_Songs)=0 then
				currentpage= totalPut \ MaxPerPage_Songs
			else
				currentpage= totalPut \ MaxPerPage_Songs + 1
			end if
		end if
		if currentPage=1 then
			showpage totalput,MaxPerPage_Songs,"default.asp"
			showsongs
		else
			if (currentPage-1)*MaxPerPage_Songs<totalPut then
				rs.move  (currentPage-1)*MaxPerPage_Songs
				dim bookmark
				bookmark=rs.bookmark
				showpage totalput,MaxPerPage_Songs,"default.asp"
				showsongs
			else
				currentPage=1
				showpage totalput,MaxPerPage_Songs,"default.asp"
				showsongs
			end if
		end if
	end if
	rs.close
%>
			</table>
		</td>
	</tr>
</table>
<%end sub %>

<%''------------------------------------ȫ�����б��ҳ'------------------------------------%>
<%
sub showsongs
	dim i
	i=0
	%>
		<tr>
			<td width="100%">��ǰλ�ã�<%=titlename%></td>
		</tr>
		<tr>
			<td>
				<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
					<tr align="center">
						<td width="35%" height="22" background="images/dh-di.jpg">��������</td>
						<td width="35%" background="images/dh-di.jpg">����ר��</td>
						<%if Switch_Comment=1 then
							response.write "<td width=8% background=images/dh-di.jpg>����</td>"
						end if%>
						<td width="15%" background="images/dh-di.jpg">��������</td>
						<td width="5%"  background="images/dh-di.jpg">��</td>
						<td width="5%"  background="images/dh-di.jpg">��</td>
						<td width="5%"  background="images/dh-di.jpg">��</td>
						<%if Session.Contents("IsAdmin")=true then%>
							<td width="5%" background="images/dh-di.jpg">ɾ</td>
							<td width="5%" background="images/dh-di.jpg">��</td>
							<td width="5%" background="images/dh-di.jpg">��</td>
						<%end if%>
					</tr>
					<%do while not rs.eof
					%>
					<tr align="center">
						<td align="left">
							<%if rs("songhot")>=5 then
								response.write "<font color=red>��</font>"
							end if
							if rs("songfile")<>"" then
								response.write (i+1)&". "&rs("songname")
							else
								response.write (i+1)&". "&rs("songname")&"(ȱ)"
							end if%>
						</td>
						<td>
						<%if rs("albumid")=0 then
							response.write "��Ա�ϴ�"
						else
							response.write "<a href='default.asp?albumid="&rs("albumid")&"'>"&rs("albumname")&"</a>"
						end if
						%>
						</td>
						<%if Switch_Comment=1 then
							response.write "<td><a href='comment_show.asp?commentindex="&rs("songid")&"'>�鿴</td>"
						end if%>
						<td><%=rs("songdatetime")%></td>
						<td><a href="lrc_show.asp?song_id=<%=rs("song_id")%>" target="_blank">
							<%if rs("songlrc")<>"" then
								response.write "<img src=images/lrc.gif border=0 alt=���>"
							else
								response.write "��"
							end if%>
							</a>
						</td>
						<%if rs("songfile")<>"" then%>
							<%if Mode_Download=1 then%>
								<td><a href=song_download.asp?song_id=<%=rs("song_id")%>><img src="images/download.gif" border="0" alt="����"></a></td>
							<%else%>
								<td><a href="<%=rs("songfile")%>"><img src="images/download.gif" border="0" alt="����"></a></td>
							<%end if%>

							<%select case Mode_Play
								case 1
							%>
								<td><a href="javascript:playonline('play1.asp?song_id=<%=rs("song_id")%>',388,145)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>
							<% case 2%>
								<td><a href="javascript:playonline('play2.asp?song_id=<%=rs("song_id")%>',380,400)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>
							<% case else%>
								<td><a href="javascript:playonline('play1.asp?song_id=<%=rs("song_id")%>',388,145)"><img src="images/play.gif" border="0" alt="���߲���"></a></td>									
							<%end select%>
						<%else%>
							<td>��</td>
							<td>��</td>
						<%end if%>

						<%if Session.Contents("IsAdmin")=true then%>
							<td><a href="song_del.asp?song_id=<%=rs("song_id")%>"><img src="images/del.gif" border="0" alt="ɾ"></a></td>
							<td><a href="song_best.asp?song_id=<%=rs("song_id")%>"><img src="images/best.gif" border="0" alt="�Ƽ�"></a></td>
							<td><a href="song_edit.asp?song_id=<%=rs("song_id")%>" target="_blank"><img src="images/edit.gif" border="0" alt="��"></a></td>
						<%end if%>
					</tr>
					<%
					i=i+1
					if i>=MaxPerPage_Songs then exit do
					rs.movenext
					loop
					%>
				</table>
			</td>
		</tr>
<% end sub %>

<%'------------------------------------��ҳ��Ϣ'------------------------------------%>
<%
function showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?keyword="&keyword&"&classid="&classid&">"
  	response.write "<center>��<font color='white'><b>"&totalnumber&"</b></font>�ף�"
  	if CurrentPage<2 then
		response.write "��ҳ&nbsp;"
		response.write "��ҳ&nbsp;"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page=1&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&CurrentPage-1&"&classid="&classid&">��ҳ</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "��ҳ&nbsp;"
		response.write "βҳ"
  	else
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&(CurrentPage+1)&"&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?keyword="&keyword&"&page="&n&"&classid="&classid&">βҳ</a>"
  	end if
   response.write "��<strong><font color=white>"&CurrentPage&"/"&n&"</font></strong>ҳ"
	%>
	ת:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
	<%next%>   
	</select>
	</form>
<% end function %>

<%'------------------------------------�û��б�'------------------------------------%>
<%sub user_list%>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>��ʱû�еǼǻ�Ա</p>"
	else
		totalPut=rs.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage_User>totalput then
			if (totalPut mod MaxPerPage_User)=0 then
				currentpage= totalPut \ MaxPerPage_User
			else
				currentpage= totalPut \ MaxPerPage_User + 1
			end if
		end if
		if currentPage=1 then
			user_showpage totalput,MaxPerPage_User,"default.asp"
			user_show
		else
			if (currentPage-1)*MaxPerPage_User<totalPut then
				rs.move  (currentPage-1)*MaxPerPage_User
				dim bookmark
				bookmark=rs.bookmark
				user_showpage totalput,MaxPerPage_User,"default.asp"
				user_show
			else
				currentPage=1
				user_showpage totalput,MaxPerPage_User,"default.asp"
				user_show
			end if
		end if
	end if
	rs.close
%>
			</table>
		</td>
	</tr>
</table>

<%end sub %>

<%''------------------------------------�û��б��ҳ'------------------------------------%>
<%
sub user_show
	dim i 
	i=0 
	%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
		<tr>
			<td width="100%"> ����ע���Ա </td>
		</tr>
		<tr>
			<td>
				<table border="1" width="100%" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#CCCCCC" bordercolorlight="#000000">
					<tr align="center">
						<td width="10%" height="22" background="images/dh-di.jpg">��Ա��</td>
						<td width="15%" background="images/dh-di.jpg">��Ա</td>
						<td width="10%" background="images/dh-di.jpg">����</td>
						<td width="20%" background="images/dh-di.jpg">���</td>
						<%if Session.Contents("KEY")="super" then%>
							<td width="25%" background="images/dh-di.jpg">IP</td>
							<td width="10%" background="images/dh-di.jpg">����</td>
							<td width="10%" background="images/dh-di.jpg">ɾ��</td>
						<%end if%>
					</tr>
					<%do while not rs.eof%>
					<tr align="center">
						<td><%=rs("userid")%></td>
						<td><a href="javascript:openScript('user_show.asp?username=<%=rs("username")%>',400,220)"><%=rs("username")%></a></td>
						<td><%=rs("userscore")%></td>
						<td>
							<%
							select case rs("userrank")
								case "super"
									userrank = "��������Ա"
								case "check"
									userrank = "�߼�����Ա"
								case "input"
									userrank = "��������Ա"
								case "user"
									userrank = "��ͨ��Ա"
								end select
							%>
							<%=userrank%>
						</td>
						<%if Session.Contents("KEY")="super" then%>
							<td><%=rs("userip")%></td>
							<td><%=rs("userpassword")%></td>
							<td><a href="admin_user_del.asp?id=<%=rs("userid")%>&amp;name=del">ɾ��</a></td>
						<%end if%>
					</tr>
					<%
					i=i+1
					if i>=MaxPerPage_User then exit do
					rs.movenext
					loop
					%>
				</table>
			</td>
		</tr>
	</table>
<% end sub %>

<%''------------------------------------�û��б��ҳ��Ϣ'------------------------------------%>
<%
function user_showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?classid="&classid&">"
  	response.write "<center>�Ǽǻ�Ա��<font color='white'><b>"&totalnumber&"</b></font>λ&nbsp;��&nbsp;"
  	if CurrentPage<2 then
		response.write "��ҳ&nbsp;&nbsp;"
		response.write "��һҳ&nbsp;&nbsp;"
  	else
		response.write "<a href="&filename&"?page=1&classid="&classid&">��ҳ</a>&nbsp;&nbsp;"
		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&classid&">��һҳ</a>&nbsp;&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "��һҳ&nbsp;&nbsp;"
		response.write "βҳ"
  	else
		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&classid&">��һҳ</a>&nbsp;&nbsp;"
		response.write "<a href="&filename&"?page="&n&"&classid="&classid&">βҳ</a>"
  	end if
   	response.write "&nbsp;��&nbsp;<strong><font color=white>"&CurrentPage&"/"&n&"</font></strong>ҳ"
	%>
	ת��:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
	<%next%>
	</select>
	</form>
<% end function %>

<%'------------------------------------����ר��'------------------------------------%>
<%sub newalbum_list%>
<table cellSpacing=0 cellPadding=0 width=403 border=0>
	<tr>
		<td vAlign=top width=403 background=<%=cenbar%> height=33></td>
	</tr>
	<tr>
		<td align=center>
			<table height=81 cellSpacing=0 cellPadding=3 width="97%" border=0>
<%	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'>��ʱ��¼ר��</p>"
	else
		totalPut=rs.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage_NewAlbum>totalput then
			if (totalPut mod MaxPerPage_NewAlbum)=0 then
				currentpage= totalPut \ MaxPerPage_NewAlbum
			else
				currentpage= totalPut \ MaxPerPage_NewAlbum + 1
			end if
		end if
		if currentPage=1 then
			newalbum_showpage totalput,MaxPerPage_NewAlbum,"default.asp"
			newalbum_show
		else
			if (currentPage-1)*MaxPerPage_NewAlbum<totalPut then
				rs.move  (currentPage-1)*MaxPerPage_NewAlbum
				dim bookmark
				bookmark=rs.bookmark
				newalbum_showpage totalput,MaxPerPage_NewAlbum,"default.asp"
				newalbum_show
			else
				currentPage=1
				newalbum_showpage totalput,MaxPerPage_NewAlbum,"default.asp"
				newalbum_show
			end if
		end if
	end if
	rs.close
%>
			</table>
		</td>
	</tr>
</table>

<%end sub %>

<%'------------------------------------����ר����ҳ'------------------------------------%>
<%sub newalbum_show%>
	<%
	i=0
	do while not rs.eof%>
		<tr>
			<td align="center">
			<%show_cd(rs("albumcover"))%><br>
			</td>
			<td width=150 align=left>
				<font color="white">
				<a href="default.asp?albumid=<%=rs("albumid")%>"><%=rs("albumname")%></a><br><br>
				���ʱ�䣺<br>
				<%=rs("albumdatetime")%>
				</font>
			</td>
		</tr>
		<%
		i=i+1
		if i>=MaxPerPage_NewAlbum then exit do
		rs.movenext
		loop
		%>
<% end sub %>

<%'------------------------------------����ר����ҳ��Ϣ'------------------------------------%>
<%
function newalbum_showpage(totalnumber,maxperpage,filename)
  	dim n
  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1
  	end if
  	response.write "<form method=Post action="&filename&"?classid="&classid&">"
  	response.write "<center>ר����<font color='white'><b>"&totalnumber&"</b></font>�ţ�"
  	if CurrentPage<2 then
		response.write "��ҳ&nbsp;"
		response.write "��ҳ&nbsp;"
  	else
		response.write "<a href="&filename&"?page=1&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?page="&CurrentPage-1&"&classid="&classid&">��ҳ</a>&nbsp;"
  	end if
  	if n-currentpage<1 then
		response.write "��ҳ&nbsp;"
		response.write "βҳ"
  	else
		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&classid="&classid&">��ҳ</a>&nbsp;"
		response.write "<a href="&filename&"?page="&n&"&classid="&classid&">βҳ</a>"
  	end if
   	response.write "��<strong><font color=white>"&CurrentPage&"/"&n&"</font></strong>ҳ"
	%>
	ת:<select name="page" size="1" style="background-color:white;FONT-SIZE: 9pt;color:black" onchange="javascript:submit()">
	<%for i = 1 to n%>   
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>   
	<%next%>   
	</select>
	</form>
<% end function %>

<%'------------------------------------ɧ���Ƽ�����ҳ�Ƽ����б�'------------------------------------%>
<%sub goodcomment_list%>
		<table cellSpacing=0 cellPadding=0 width=403 background=Images/cenbar2.gif border=0>
			<tr>
				<td>
					<table cellSpacing=0 cellPadding=0 width=403 border=0>
						<tr>
							<td vAlign=top width=403 background="images/cenbar1-4.gif" height=33>&nbsp;</td>
						</tr>
						<tr>
							<td align=center>
								<table height=81 cellSpacing=0 cellPadding=2 width="97%" border=0>
								<%
								sql="select * from goodcomment order by goodcomment_id desc;"  
								rs.Open sql,conn,1,1
								do while not(rs.eof) 
								%>
									<tr>
										<td width="38%" align=left>
											<a href="default.asp?goodcomment_id=<%=rs("goodcomment_id")%>"><% =rs("goodcommenttitle")%></a>
										</td>
									</tr>
								<%
								rs.movenext
								loop
								rs.close
								%>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
<% end sub %>

<%'------------------------------------��ʾ��CD��ʽ'------------------------------------%>
<%function show_cd(cdcover)%>
	<table border="1" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="3" style="border: 1px solid rgb(0,0,0)" height="5" width="200" bgcolor="#FFFFFF"><span style="font-size: 1pt" style="line-height: 1pt">��PengJohn's ���������</span></td>
		</tr>
		<tr>
			<td style="border-left: medium none rgb(0,0,0)" width="20" height="137" bgcolor="#000000" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF"></td>
			<td style="border-left: 1px solid rgb(0,0,0); border-right: 1px solid rgb(0,0,0)" width="3" height="137" bgcolor="#FFFFFF"><span style="font-size: 1pt">d</span></td>
			<td style="border-right: 3px solid rgb(0,0,0)" width="200" height="200"><img src="<%=cdcover%>"  width="200" height="200" border="0"></td>
		</tr>
		<tr>
			<td height="5" colspan="3" style="border: 1px solid rgb(0,0,0)" width="200" bgcolor="#FFFFFF"><span style="font-size: 1pt" style="line-height: 1pt">��PengJohn's ���������</span></td>
		</tr>
	</table>
<%end function%>

<%'------------------------------------ɧ���Ƽ�����ҳ�Ƽ���'------------------------------------%>
<%function goodcomment_show%>
		<table cellSpacing=0 cellPadding=0 width=403 background=Images/cenbar2.gif border=0>
			<tr>
				<td>
					<table cellSpacing=0 cellPadding=0 width=403 border=0>
						<tr>
							<td vAlign=top width=403 background="images/cenbar1-4.gif" height=33>&nbsp;</td>
						</tr>
						<tr>
							<td align=center>
								<table height=81 cellSpacing=0 cellPadding=2 width="97%" border=0>
									<tr>
										<td width="38%" align=left>
																		
											<%
											k=0
											if goodcomment_id<>"" then
												sql="select * from goodcomment where goodcomment_id="&cstr(goodcomment_id)&" order by goodcomment_id desc;"
											else
												sql="select * from goodcomment order by goodcomment_id desc;"
											end if
											rs.Open sql,conn,1,1
											do while not(rs.eof) 
												response.write "<b>"&rs("goodcommenttitle")&"</b><br><br>"& changechr(rs("goodcommentcontent")) &"<HR SIZE=1>"
												k=k+1
												if k>=Num_Goodcomment then exit do
											rs.movenext
											loop
											rs.close
											%>
											::<a href="default.asp?classid=87">�����Ƽ�</a>::<br>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td vAlign=top width="100%"><img src="Images/cenbar3.gif" border=0></td>
			</tr>
		</table>
<%end function%>


<%sub TopSongs_Show%>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar2.gif" border=0></td>
			</tr>
			<tr>
				<td width="100%" align=center>
					<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width=166 bgColor=#055192 border=0>
						<tr>
							<td width="100%" align=center>
							<table cellPadding=2 width="100%" border=0>
							<%
									k=0
									sql="select * from song order by songhit desc;"  
									rs.Open sql,conn,1,1
									do while not(rs.eof) 							
							%>
								<tr>
									<td class=jnfont4 width="90%"><%=rs("songname")%></td>
									<td class=jnfont4 width="10%"><%=rs("songhit")%></td>
								</tr>
								<%
									k=k+1
									if k>=Num_TopSong then exit do
									rs.movenext
									loop
									rs.close
								%>
							</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar1-2.gif" border=0></td>
			</tr>
<%end sub%>

<%sub TopAlbum_Show%>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar3.gif" border=0></td>
			</tr>
			<tr>
				<td width="100%">
					<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width=166 align=center bgColor=#055192 border=0>
						<tr>
							<td width="100%">
								<table cellPadding=2 width="100%" border=0>
									<tr>
										<td class=jnfont4 width="90%"></td>
										<td class=jnfont4 width="10%"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar1-2.gif" border=0></td>
			</tr>
<%end sub%>

<%sub TotalCount_Show%>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar1.gif" border=0></td>
			</tr>
			<tr>
				<td width="100%" class=jnfont4>
					<table style="BORDER-RIGHT: #0563bb 2px solid; BORDER-LEFT: #0563bb 2px solid; BORDER-COLLAPSE: collapse" cellPadding=3 width=166 align=center bgColor=#055192 border=0>
						<tr>
							<td vAlign=top width="100%">
									<%
									response.write "���������У�<br>"
									sql="select * from singer"
									rs.Open sql,conn,1,1
										singertotalcount =INT(RS.recordcount)
									rs.close

									sql="select * from album where (albumid mod 100<>0)"
										rs.Open sql,conn,1,1
									albumtotalcount =INT(RS.recordcount)
									rs.close

									sql="select * from song where songfile<>''"
									rs.Open sql,conn,1,1
										songtotalcount =INT(RS.recordcount)
									rs.close
									%>
									����������<%=singertotalcount%><BR>
									ר��������<%=albumtotalcount%><BR>
									����������<%=songtotalcount%><BR><br>
									<!--#include file="count_show.asp"-->					
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align=middle width="100%"><img src="Images/rightbar1-2.gif" border=0></td>
			</tr>
<%end sub%>			