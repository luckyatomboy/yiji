<%
if session("redboy_username")="" then
%>
<script language="javascript">
top.location.href="../index.asp"
</script>
<%  
  response.end
end if
%>
<!-- #include file="../conn2.asp" -->
<!-- #include file="../const.asp" -->

<html>
<head>
<title><%=dianming%> - ���ڱ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla1="0" and fla2="0" then
%>	
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<script language="javascript">
function query_shipment()
{
	window.location.href="query_shipment.asp"
}
function query_shipment_new()
{
	window.location.href="query_shipment_new.asp"
}
function add_shipment()
{
	window.location.href="add_shipment.asp"
}
function add_shipment_new()
{
	window.location.href="add_shipment_new.asp"
}
function create_stock_order()
{
	window.location.href="create_stock_order.asp"
}
function create_stock_order_new()
{
	window.location.href="create_stock_order_new.asp"
}
function query_stock_order()
{
	window.location.href="query_stock_order.asp"
}
function query_stock_order_new()
{
	window.location.href="query_stock_order_new.asp"
}
function create_stock_issue()
{
	window.location.href="create_stock_issue.asp"
}
function query_stock_issue()
{
	window.location.href="query_stock_issue.asp"
}
function upload_excel()
{
	window.location.href="upload_excel.asp"
}
function download_excel()
{
	window.location.href="download_excel.asp"
}
function create_sales_contract()
{
	window.location.href="create_sales_contract.asp"
}
</script>
<form name="form1">
<!-- ----------------ά�����ڱ�����------------- -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;���ڱ�</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>

<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr>
    <td width="25%" height="30" align="right">
    </td>
    <td width="75%" class="category">
		  <input type="button" value=" �������ڱ� " onClick="return add_shipment()">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="button" value=" �������ڱ�(��) " onClick="return add_shipment_new()">
		</td>
  </tr>
  <tr>
  	<td></td>
		<td>
		  <input type="button" value=" ���ڱ��ѯ " onClick="return query_shipment()">&nbsp;&nbsp;&nbsp;&nbsp;		
		  <input type="button" value=" ���ڱ��ѯ(��) " onClick="return query_shipment_new()">		
		</td>        	
	</tr>
  <tr>
  	<td></td>
		<td>
		  <input type="button" value=" ���ڱ���" onClick="return upload_excel()">&nbsp;&nbsp;&nbsp;&nbsp;		
		  <input type="button" value=" ���ڱ���" onClick="return download_excel()">		
		</td>        	
	</tr>	
  <tr>
  	<td></td>
		<td>
		  <input type="button" value=" ����������ͬ" onClick="return create_sales_contract()">	
		</td>        	
	</tr>		
</table>
</td>
<td></td>
</tr>

<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;������</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>

<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr>
    <td width="25%" height="30" align="right">
    </td>
    <td width="75%" class="category">
		  <input type="button" value=" ������ⵥ " onClick="return create_stock_order()">&nbsp;&nbsp;&nbsp;&nbsp;	
		  <input type="button" value=" ������ⵥ(��) " onClick="return create_stock_order_new()">
		</td>
  </tr>
  <tr>
  	<td></td>
		<td>
		  <input type="button" value=" ��ⵥ��ѯ " onClick="return query_stock_order()">	&nbsp;&nbsp;&nbsp;&nbsp;	
		  <input type="button" value=" ��ⵥ��ѯ(��) " onClick="return query_stock_order_new()">	
		</td>        	
	</tr>
</table>
</td>
<td></td>
</tr>

<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�������</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>

<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr>
    <td width="25%" height="30" align="right">
    </td>
    <td width="75%" class="category">
		  <input type="button" value=" �������ⵥ " onClick="return create_stock_issue()">
		</td>
  </tr>
  <tr>
  	<td></td>
		<td>
		  <input type="button" value=" ���ⵥ��ѯ " onClick="return query_stock_issue()">		
		</td>        	
	</tr>
</table>
</td>
<td></td>
</tr>

<tr>
<td><img src="../images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../images/r_3.gif" alt="" /></td>
</tr>
</table>


</form>

</body>
</html>