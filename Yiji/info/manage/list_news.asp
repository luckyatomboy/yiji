<%@language="VBscript"%>
<%Response.Expires=0%>
<%dim ThisKey
ThisKey = "d"
%>
<LINK href="../Image/style.css" type=text/css rel=stylesheet>
<link href="../../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>

<!--#include file="../include/checkadmin.asp"-->
<!--#include file="../include/conn.asp"-->
<!-- #include file="../../const.asp" --> 
<%
if fla89="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<style>
	td{font-size:13px;}
</style>
<script language="javascript">
	//�޸�
	function doEdit(){
		if(checkedNum()>0){
			document.frmlist.action="edit_news.asp";
			document.frmlist.submit();
		}else{
			alert("��ѡ����Ҫ�޸ĵ����ţ�");
		}
	}
	//ɾ��
	function doDelete(){
		if(checkedNum()>0){
			if(confirm("ȷ��Ҫɾ����ѡ���������\nע�⣺ɾ�����ָܻ���")){
				document.frmlist.action="delete_news.asp";
				document.frmlist.submit();
			}
		}else{
			alert("��ѡ����Ҫɾ�������ţ�");
		}
	}
	//ѡ��
	function doSelect(){
		if(document.frmlist.btnSelect.value=="ȫѡ"){
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=true;
			
			document.frmlist.btnSelect.value="ȡ��";
			//document.frmlist.checkall.checked=true;
		}else{
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=false;
			document.frmlist.btnSelect.value="ȫѡ";
			//document.frmlist.checkall.checked=false;
		}
	}
	//ȫѡ
	function doCheckall(){
		if(document.frmlist.id[0].checked){
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=true;
			
			document.frmlist.btnSelect.value="ȡ��";
		}else{
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=false;
			document.frmlist.btnSelect.value="ȫѡ";
		}
	}
	//��Ŀ
	function checkedNum(){
		var num=0;
		for(var i=0;i<document.frmlist.id.length;i++){
			if(document.frmlist.id[i].checked)
				num++;
		}
		return num;
	}
</script>
<%
keyword=request.form("keyword")
days=request.form("days")
columnid=request.form("columnid")
if columnid="" or isnull(columnid) or not isnumeric(columnid) then
	columnid=0
else
	columnid=cint(columnid)
end if

if days="" or isnull(days) or not isnumeric(days) then days="0"
if keyword<>"" then search_string = search_string & " and (title like '%"&keyword&"%' or content like '%"&keyword&"%')"
if days<>"" and days<>"0" then search_string = search_string & " and datediff('d',regdate,now())<"&days
if columnid<>0 then search_string = search_string & " and columnid="&columnid
%>
<table  class="toptable grid" border="1" cellspacing="1" width="600" >
  <tr>
    <td height="23" align="center" background="../../images/r_1.gif"><img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_news.asp"><b>ά������</b></a></font>
  <img src="../../images/r_0.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news.asp"><b>��������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news_catalog.asp"><b>�������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_review.asp"><b>���۹���</b></a></font></td>
  </tr>
  <tr>
    <td height="23" bgcolor="#FFFFFF">
	<!--search-->
	<br>
	<div align="center">
	<form name="frmsearch" method="post">
	�ؼ��֣�
	<input type="text" name="keyword" value="<%=keyword%>">
	<select name="columnid">
	<option value="0">����Ŀ¼</option>
	<%
		set rs_root=conn.execute("select columnid,columnname from tbl_news_column where columnparentid=0 order by columnid desc")
		do while not rs_root.eof
			if columnid=rs_root("columnid") then
				response.write("<option selected value="&rs_root("columnid")&">��"&rs_root("columnname")&"</option>")
			else
				response.write("<option value="&rs_root("columnid")&">��"&rs_root("columnname")&"</option>")
			end if
			set rs_child=conn.execute("select columnid,columnname from tbl_news_column where columnparentid="&rs_root("columnid")&" order by columnid desc")
			do while not rs_child.eof
				if columnid=rs_child("columnid") then
					response.write("<option selected value="&rs_child("columnid")&"> ����"&rs_child("columnname")&"</option>")
				else
					response.write("<option value="&rs_child("columnid")&"> ����"&rs_child("columnname")&"</option>")
				end if
				rs_child.movenext
			loop
			rs_root.movenext
		loop
		set rs_root=nothing
	%>
	</select>
	<select name="days">
	<option value="0" <%if days="0" then response.write("selected")%>>����</option>
	<option value="1" <%if days="1" then response.write("selected")%>>����</option>
	<option value="10" <%if days="10" then response.write("selected")%>>��10��</option>
	<option value="20" <%if days="20" then response.write("selected")%>>��20��</option>
	<option value="30" <%if days="30" then response.write("selected")%>>��30��</option>
	</select>
	<input type="submit" name="btnSubmit" value="����" style="width:60px;">
	</form>
	</div>

	<!--list-->
	<table align="center"  class="toptable grid" border="1" width="96%" cellpadding="1" cellspacing="1" >
	  <form name="frmlist" method="post">
	  
	  <tr bgcolor="#FFFFFF" align="center">
		<td width="30" class=forumrow><b>ID</b></td>
		<td width="30" class=forumrow><input type="checkbox" onclick="javascript:doCheckall();" name="id" value="0"></td>
		<td class=forumrow><b>����</b></td>
		<td  class=forumrow width="60"><b>�����</b></td>
		<td  class=forumrow width="80"><b>��������</b></td>
		<td  class=forumrow width="40"><b>��ҳ</b></td>
		<td  class=forumrow width="40"><b>�ص�</b></td>
	  </tr>
	  <%
	  dim sqlstr
	  dim intPage,intRecordcount,intPagecount,itotalpage

	  intPage=request.form("intPage")
	  if intPage="" or isnull(intPage) or not isnumeric(intPage) then intPage=1
	  intPage=cint(intPage)

	  sqlstr="select id,title,imgname,hits,regdate,index,important from tbl_news order by id desc"
	  if search_string<>"" then
		sqlstr="select id,title,imgname,hits,regdate,index,important from tbl_news where 1<>0 "&search_string&" order by id desc"
	  end if

	  set rs=server.createobject("adodb.recordset")
	  rs.open sqlstr,conn,3,1
	  intRecordcount=rs.recordcount
	  if intRecordcount>0 then
		  rs.pagesize=8
		  intPagecount=rs.pagecount
		  if intPage>intPagecount or intPage<1 then intPage=1
		  rs.absolutePage=intPage
		  for i=1 to rs.pagesize
	  %>
	  <tr bgcolor="#FFFFFF" align="center">
		<td height="26"><%=rs("id")%></td>
		<td><input type="checkbox" name="id" value="<%=rs("id")%>"></td>
		<td align="left">
		<a href="../docc/news_detail.asp?id=<%=rs("id")%>"><%=rs("title")%></a>
		<%if rs("imgname")<>"" then response.write("��ͼ��")%>
		</td>
		<td><%=rs("hits")%></td>
		<td><%=formatdatetime(rs("regdate"),2)%></td>
		<td><%if rs("index")="1" then%><img src="../images/ok.gif"><%end if%></td>
		<td><%if rs("important")="1" then%><img src="../images/ok.gif"><%end if%></td>
	  </tr>
	  <%
			rs.movenext
			if rs.eof then exit for
		  next
	  end if
	  %>
	  <tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" height="60">
		<input type="button" name="btnSelect" value="ȫѡ" onclick="javascript:doSelect();" style="width:60px;">
		<input type="button" name="bntEdit" value="�޸�" onclick="javascript:doEdit();" style="width:60px;">
		<input type="button" name="bntDelete" value="ɾ��" onclick="javascript:doDelete();" style="width:60px;">
		</td>
	  </tr>
	  <input type="hidden" name="intPage" value="<%=intPage%>">
	  <input type="hidden" name="keyword" value="<%=keyword%>">
	  <input type="hidden" name="days" value="<%=days%>">
	  <input type="hidden" name="columnid" value="<%=columnid%>">
	  </form>
	</table>
	<br>
	</td>
  </tr>
  <!--showpage-->
  <form name="frmpage" method="post">
  <input type="hidden" name="keyword" value="<%=keyword%>">
  <input type="hidden" name="days" value="<%=days%>">
  <input type="hidden" name="columnid" value="<%=columnid%>">
  <tr>
    <td colspan="2" height="23" align="center">
	<!--#include file="../include/showpage.asp"-->
	</td>
  </tr>
  </form>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../../images/r_1.gif" alt="" /></td>
<td width="100%" background="../../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>��</td>
	  <td align="right">��</td>
    
    </tr>
  </table>
  </table>
