<%@language="VBscript"%>
<%Response.Expires=0%>
<link href="../../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style><%dim ThisKey
ThisKey = "d"
%>

<!--#include file="../include/checkadmin.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/tools.asp"-->
<style>
	td{font-size:13px;}
</style>
<script language="javascript">
	//ɾ��
	function doDelete(){
		if(checkedNum()>0){
			if(confirm("ȷ��Ҫɾ����ѡ���������\nע�⣺ɾ�����ָܻ���")){
				document.frmlist.action="delete_review.asp";
				document.frmlist.submit();
			}
		}else{
			alert("��ѡ����Ҫɾ�������ۣ�");
		}
	}
	//ѡ��
	function doSelect(){
		if(document.frmlist.btnSelect.value=="ȫѡ"){
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=true;
			
			document.frmlist.btnSelect.value="ȡ��";
		}else{
			for(var i=0;i<document.frmlist.id.length;i++)
				document.frmlist.id[i].checked=false;
			document.frmlist.btnSelect.value="ȫѡ";
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

if days="" or isnull(days) or not isnumeric(days) then days="0"
if keyword<>"" then search_string = search_string & " and content like '%"&keyword&"%'"
if days<>"" and days<>"0" then search_string = search_string & " and datediff('d',regdate,now())<"&days
%>
<table border="0" cellspacing="1" width="100%" >
  <tr>
    <td height="23" align="center"background="../../images/r_1.gif"><img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_news.asp"><b>ά������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news.asp"><b>��������</b></a></font>
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
	<table align="center" border="0" width="96%" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
	  <form name="frmlist" method="post">
	  <tr bgcolor="#FFFFFF" align="center">
		<td width="30"><b>ID</b></td>
		<td width="30"><input type="checkbox" onclick="javascript:doCheckall();" name="id" value="0"></td>
		<td><b>����</b></td>
	  </tr>
	  <%
	  dim sqlstr
	  dim intPage,itotalrecord,intPagesize,itotalpage

	  intPage=request.form("intPage")
	  if intPage="" or isnull(intPage) or not isnumeric(intPage) then intPage=1
	  intPage=cint(intPage)

	  sqlstr="select * from V_News_Review where 1<>0 "&search_string&" order by reviewid desc"

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
		<td height="26"><%=rs("reviewid")%></td>
		<td><input type="checkbox" name="id" value="<%=rs("reviewid")%>"></td>
		<td align="left">
		[<%=rs("regdate")%> <a href="mailto:<%=txt2html(rs("yourmail"))%>"><%=txt2html(rs("yourname"))%></a> �� <a target="_blank" href="../docc/news_detail.asp?id=<%=rs("newsid")%>"><%=txt2html(rs("title"))%></a> ����]<BR><%=txt2html(rs("content"))%> 
		</td>
	  </tr>
	  <%
			rs.movenext
			if rs.eof then exit for
		  next
	  end if
	  %>
	  <tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" height="60">
		<input type="button" name="btnSelect" value="ȫѡ" onclick="javascript:doSelect();" style="width:60px;">
		<input type="button" name="bntDelete" value="ɾ��" onclick="javascript:doDelete();" style="width:60px;">
		</td>
	  </tr>
	  <input type="hidden" name="intPage" value="<%=intPage%>">
	  <input type="hidden" name="keyword" value="<%=keyword%>">
	  <input type="hidden" name="days" value="<%=days%>">
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
