<%@language="VBscript"%>
<%Response.Expires=0%>
<link href="../../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<%dim ThisKey
ThisKey = "d"
%>

<!--#include file="../include/checkadmin.asp"-->
<!--#include file="../include/conn.asp"-->
<style>
	td{font-size:13px;}
</style>
<table border="0" cellspacing="1" width="90%" >
  <tr>
    <td height="23" align="center"background="../../images/r_1.gif"><img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_news.asp"><b>ά������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news.asp"><b>��������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="add_news_catalog.asp"><b>�������</b></a></font>
  <img src="images/spacer.gif" width="10" height="1"><font style="font-size:13px;" ><a href="list_review.asp"><b>���۹���</b></a></font></td>
  </tr>
  <tr>
    <td height="23" bgcolor="#FFFFFF">
	<br>
	<table border="0" width="100%">
	  <tr>
		<td width="50%" valign="top">
		<!--tree start-->
		<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" bgcolor="#D6DFF7">
		  <tr>
			<td height="23">����Ŀ¼�б�</td>
		  </tr>
		  <tr>
			<td bgcolor="#FFFFFF">
			<br>
			<!--#include file="../include/news_tree.asp"-->
			<br>
			</td>
		  </tr>
		</table>
		<!--tree end-->
		</td>
		<td width="50%"  valign="top">
		<!--add start-->
		<script language="javascript">
			function doAdd(){
				if(document.frmadd.columnname.value==""){
					alert("������Ŀ¼���ƣ�");
					document.frmadd.columnname.focus();
					return false;
				}
			}
		</script>
		<form name="frmadd" method="post" action="add_news_catalog_ok.asp" onsubmit="javascript:return doAdd();">
		<input type="hidden" name="action" value="addnew">
		<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" bgcolor="#D6DFF7">
		  <tr>
			<td height="23">����Ŀ¼</td>
		  </tr>
		  <tr bgcolor="#FFFFFF">
			<td>
			<table border="0" width="100%">
			  <tr>
				<td height="28" width="60">&nbsp; Ŀ¼��</td>
				<td><input type="text" name="columnname"></td>
			  </tr>
			  <tr>
				<td height="28" width="60">&nbsp; ���ڣ�</td>
				<td>
				<select name="columnid">
				<option value="0">����Ŀ¼</option>
				<%
				set rs_root=conn.execute("select columnid,columnname from tbl_news_column where columnparentid=0 order by columnid desc")
				do while not rs_root.eof
					response.write("<option value="&rs_root("columnid")&">"&rs_root("columnname")&"</option>")
					rs_root.movenext
				loop
				set rs_root=nothing
				%>
				</select>
				</td>
			  </tr>
			  <tr>
				<td height="28"></td>
				<td>
				<input type="submit" name="btnSubmit" value="����" style="width:60px;">
				<input type="reset" name="btnReset" value="ȡ��" style="width:60px;">
				</td>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
		</form>
		<!--add end-->
		<br>
		<!--edit start-->
		<script language="javascript">
			function doEdit(){
				if(document.frmedit.columnid.value==""){
					alert("��ѡ����Ҫ�޸ĵ�Ŀ¼��");
					return false;
				}
				if(document.frmedit.columnname.value==""){
					alert("�������µ�Ŀ¼���ƣ�");
					document.frmedit.columnname.focus();
					return false;
				}
			}
		</script>
		<form name="frmedit" action="add_news_catalog_ok.asp" method="post" onsubmit="javascript:return doEdit();">
		<input type="hidden" name="action" value="edit">
		<input type="hidden" name="columnid" value="">
		<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" bgcolor="#D6DFF7">
		  <tr>
			<td height="23">�޸�Ŀ¼</td>
		  </tr>
		  <tr bgcolor="#FFFFFF">
			<td>
			<table border="0" width="100%">
			  <tr>
				<td height="28" width="60">&nbsp; ԭ����</td>
				<td><input type="text" readonly name="old_columnname"></td>
			  </tr>
			  <tr>
				<td height="28">&nbsp; ������</td>
				<td><input type="text" name="columnname"></td>
			  </tr>
			  <tr>
				<td height="28"></td>
				<td>
				<input type="submit" name="btnSubmit" value="ȷ��" style="width:60px;">
				<input type="reset" name="btnReset" value="ȡ��" style="width:60px;">
				</td>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
		</form>
		<!--edit end-->
		</td>
	  </tr>
	</table>
	<br>
	</td>
  </tr>
  <tr>
    <td colspan="2" height="23" align="center"></td>
  </tr>
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
