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
<title><%=dianming%> - ���ڱ��ѯ</title>
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
if fla1="0" and fla2="0" and fla3="0" and fla4="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%
'ȡ�������ؼ���  
nowkeyword=request("keyword") 
nowsales=request("sales")
nowcustomer=request("customer")
nowfrom=request("datefrom")
nowto=request("dateto")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="form2">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
<!--
  <input type="hidden" name="field4" value="<%=request("field4")%>">  
  <input type="hidden" name="field5" value="<%=request("field5")%>">
  <input type="hidden" name="field6" value="<%=request("field6")%>">
  <input type="hidden" name="field7" value="<%=request("field7")%>">
  <input type="hidden" name="field8" value="<%=request("field8")%>">    
  <input type="hidden" name="field9" value="<%=request("field9")%>">    
-->
  <input type="hidden" name="span1" value="<%=request("span1")%>">
  <tr> 
<!--  <td width="700" height="21" align="right"><font size=4><b>��ѯ���ڱ�</b></font></td> -->
	<td align="right">
	  ������
		<%
			sql="select * from sales order by name"
			set rs_sales=conn.execute(sql)
		%>
			<select name="sales" onChange="form2.submit()">
		     <option value="">��������Ա</option>
	  <%
			do while rs_sales.eof=false
		%>
		    <option value="<%=rs_sales("name")%>"<%if trim(cstr(rs_sales("name")))=nowsales then%> selected="selected"<%end if%>><%=rs_sales("name")%></option>
		<%
			  rs_sales.movenext
			loop
		%>
		  </select>	  

		<%
			sql="select * from customer order by customername"
			set rs_customer=conn.execute(sql)
		%>
			<select name="customer" onChange="form2.submit()">
		     <option value="">���пͻ�</option>
	  <%
			do while rs_customer.eof=false
		%>
		    <option value="<%=rs_customer("customername")%>"<%if trim(cstr(rs_customer("customername")))=nowcustomer then%> selected="selected"<%end if%>><%=rs_customer("customername")%></option>
		<%
			  rs_customer.movenext
			loop
		%>
		  </select>	  	  
		<input type="text" name="keyword" size="20" value="<%=nowkeyword%>">
<!--	  <input type="submit" value=" ��ѯ " class="button">&nbsp; -->
	</td>
  </tr>
  <tr> 
	<td align="right">
	  �������ڣ���
	  <input name="datefrom" style="width:80px" onClick="JavaScript:window.open('../day.asp?form=form2&field=datefrom&oldDate='+datefrom.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
	  &nbsp;&nbsp;&nbsp;&nbsp;��
	  <input name="dateto" style="width:80px" onClick="JavaScript:window.open('../day.asp?form=form2&field=dateto&oldDate='+dateto.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">
	  <input type="submit" value=" ��ѯ " class="button">&nbsp;
	</td>
  </tr>  
</form>  
</table>
<%
  sql="select * from shipment where shipmentnum<=99999999" <!-- where shipmentnum like '1%' or shipmentnum like '2%'" -->
  if nowsales<>"" then
  	sql=sql&" and sales='"&nowsales&"'"
  end if
  if nowcustomer<>"" then
  	sql=sql&" and customer='"&nowcustomer&"'"
  end if  
  if nowkeyword<>"" then
  	sql=sql&" and shipmentnum="&nowkeyword
  end if  
  if nowfrom<>"" and nowto="" then
	sql=sql&" and boarddate=#"&nowfrom&"#"
  elseif nowfrom="" and nowto<>"" then
	sql=sql&" and boarddate=#"&nowto&"#"
  elseif nowfrom<>"" and nowto<>"" then
	sql=sql&" and boarddate>=#"&nowfrom&"# and boarddate<=#"&nowto&"#"
  end if

	
  if request("order1")<>"" then
    sql=sql&" order by shipmentnum "&request("order1")
  elseif request("order2")<>"" then
    sql=sql&" order by status "&request("order2")
  elseif request("order3")<>"" then
    sql=sql&" order by customer "&request("order3") 
  elseif request("order4")<>"" then
    sql=sql&" order by vendor "&request("order4") 
  elseif request("order5")<>"" then
    sql=sql&" order by contract "&request("order5") 
  elseif request("order6")<>"" then
    sql=sql&" order by material "&request("order6")
  elseif request("order7")<>"" then
    sql=sql&" order by country "&request("order7")
  elseif request("order8")<>"" then
    sql=sql&" order by boarddate "&request("order8")
  elseif request("order9")<>"" then
    sql=sql&" order by plant "&request("order9") 
  elseif request("order10")<>"" then
    sql=sql&" order by deliverydate "&request("order10") 
  elseif request("order11")<>"" then
    sql=sql&" order by case "&request("order11")
  elseif request("order12")<>"" then
    sql=sql&" order by dongjiancom "&request("order12")
  elseif request("order13")<>"" then
    sql=sql&" order by zidongcom "&request("order13")
  end if
  
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;���ڱ��ѯ</td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../images/r_2.gif" alt="" /></td>
</tr>

<tr>
<td></td>
<td>
<!--startprint-->
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
<form name="form1" action="produit_del.asp">
  <input type="hidden" name="queryform" value="<%=request("queryform")%>">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
<!--
  <input type="hidden" name="field4" value="<%=request("field4")%>">
  <input type="hidden" name="field5" value="<%=request("field5")%>">
  <input type="hidden" name="field6" value="<%=request("field6")%>">
  <input type="hidden" name="field7" value="<%=request("field7")%>">
  <input type="hidden" name="field8" value="<%=request("field8")%>">
  <input type="hidden" name="field9" value="<%=request("field9")%>">
-->
  <input type="hidden" name="company" value="<%=nowcompany%>">
  <input type="hidden" name="licensetype" value="<%=nowlicensetype%>">
  <input type="hidden" name="license" value="<%=nowlicense%>">
  <input type="hidden" name="keyword" value="<%=nowkeyword%>">
  <input type="hidden" name="order1" value="<%=request("order1")%>">
  <input type="hidden" name="order2" value="<%=request("order2")%>">
  <input type="hidden" name="order3" value="<%=request("order3")%>">
  <input type="hidden" name="order4" value="<%=request("order4")%>">
  <input type="hidden" name="order5" value="<%=request("order5")%>">
  <input type="hidden" name="order6" value="<%=request("order6")%>">
  <input type="hidden" name="order7" value="<%=request("order7")%>">
  <input type="hidden" name="order8" value="<%=request("order8")%>">
  <input type="hidden" name="order9" value="<%=request("order9")%>">
  <input type="hidden" name="order10" value="<%=request("order10")%>">
  <input type="hidden" name="order11" value="<%=request("order11")%>">
  <input type="hidden" name="order12" value="<%=request("order12")%>">
  <input type="hidden" name="order13" value="<%=request("order13")%>">
  <tr align="center">
	<td class="category" width="60" height="30">
		<a href="?order1=<%if request("order1")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">���<%if request("order1")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order2=<%if request("order2")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">״̬<%if request("order2")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="100" height="30">
	  <a href="?order3=<%if request("order3")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�ͻ�����<%if request("order3")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="100" height="30">
		<a href="?order4=<%if request("order4")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��Ӧ������<%if request("order4")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="120" height="30">
		<a href="?order5=<%if request("order5")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��ͬ��<%if request("order5")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order6=<%if request("order6")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">Ʒ��<%if request("order6")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order7=<%if request("order7")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">����<%if request("order7")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order9=<%if request("order9")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">����<%if request("order9")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="100" height="30">
		<a href="?order12=<%if request("order12")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">���칫˾<%if request("order12")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
	<td class="category" width="100" height="30">
		<a href="?order13=<%if request("order13")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�Զ���˾<%if request("order13")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
	<td class="category" width="80" height="30">
		<a href="?order8=<%if request("order8")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">������<%if request("order8")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
	<td class="category" width="80" height="30">
		<a href="?order10=<%if request("order10")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��������<%if request("order10")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
	<td class="category" width="80" height="30">
		<a href="?order11=<%if request("order11")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">���<%if request("order11")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
  <td class="category">�޸�</td>
  </tr>
  <%
  set rs_shipment =server.createobject("ADODB.RecordSet")	
  rs_shipment.open sql,conn,1,3
  if not rs_shipment.eof then
  do while rs_shipment.eof=false
  %>
  <!--<tr align="center" onDblClick="window.opener.document.<%=request("queryform")%>.<%=request("field")%>.value='<%=rs_shipment("shipmentnum")%>';window.opener.document.<%=request("queryform")%>.<%=request("field2")%>.value='<%=rs_shipment("material")%>';window.opener.document.<%=request("queryform")%>.<%=request("field3")%>.value='<%=rs_shipment("country")%>';window.opener.document.<%=request("queryform")%>.<%=request("field4")%>.value='<%=rs_shipment("case")%>';window.opener.document.<%=request("queryform")%>.<%=request("field5")%>.value='<%=rs_shipment("plant")%>';window.opener.document.<%=request("queryform")%>.<%=request("field6")%>.value='<%=rs_shipment("contract")%>';window.opener.document.<%=request("queryform")%>.<%=request("field7")%>.value='<%=rs_shipment("spec")%>';window.opener.document.<%=request("queryform")%>.<%=request("field8")%>.value='<%=rs_shipment("customer")%>';window.close();" <%if rs_shipment("status")="����" then%>bgcolor="darkgrey"<%elseif rs_shipment("status")="ͨ����" then%>bgcolor="lawngreen"<%elseif rs_shipment("status")="���ͻ�" then%>bgcolor="red"<%end if%>>-->
    <tr align="center" onDblClick="window.opener.document.<%=request("queryform")%>.<%=request("field")%>.value='<%=rs_shipment("shipmentnum")%>';window.opener.document.<%=request("queryform")%>.<%=request("field2")%>.value='<%=rs_shipment("material")%>';window.opener.document.<%=request("queryform")%>.<%=request("field3")%>.value='<%=rs_shipment("plant")%>';window.close();" <%if rs_shipment("status")="����" then%>bgcolor="darkgrey"<%elseif rs_shipment("status")="ͨ����" then%>bgcolor="lawngreen"<%elseif rs_shipment("status")="���ͻ�" then%>bgcolor="red"<%end if%>>
	<td align="center" height="30"><%=rs_shipment("shipmentnum")%></td>
  <td align="center"><%=rs_shipment("status")%></td>	  
  <td align="center"><%=rs_shipment("customer")%></td>
  <td align="center"><%=rs_shipment("vendor")%></td>
  <td align="center"><%=rs_shipment("contract")%></td>
  <td align="center"><%=rs_shipment("material")%></td>
  <td align="center"><%=rs_shipment("country")%></td>
  <td align="center"><%=rs_shipment("plant")%></td>
  <td align="center"><%=rs_shipment("dongjiancom")%></td>
  <td align="center"><%=rs_shipment("zidongcom")%></td>
  <td align="center"><%=rs_shipment("boarddate")%></td>
  <td align="center"><%=rs_shipment("deliverydate")%></td>
  <td align="center"><%=rs_shipment("case")%></td>
  <td align="center">
    	<a href="change_shipment.asp?form=<%=request("form")%>&shipment=<%=rs_shipment("shipmentnum")%>&keyword=<%=nowkeyword%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">�޸�</a>
  </td>
  </tr>
  </tr>
  <%
  	'end if
    rs_shipment.movenext
  loop
  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td colspan="12" height="25" align="center" style="color:red"><b>û���ҵ���¼</b></td>
  </tr>
  <%
  end if
  %>
  <%
  if rs_shipment.recordcount>0 then 
  %> 
</table></td></tr>
<%end if%> 
</form>   
</table>
<!--endprint-->
</td>
<td></td>
</tr>

</table>

</body>
</html>