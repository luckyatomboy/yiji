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
<title><%=dianming%> - ��ѯ</title>
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
if fla34="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%
'ȡ�������ؼ���  
nowkeyword=request("keyword") 
nowrefstockdoc=request("refstockdoc") 
nowfrom=request("datefrom")
nowto=request("dateto")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="form2">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="span1" value="<%=request("span1")%>">
  <tr> 
	<td align="right">
	  ������
		<input type="text" name="keyword" size="20" value="<%=nowkeyword%>">
	</td>
  </tr>
  <tr> 
	<td align="right">
	  �ο���ⵥ��
		<input type="text" name="refstockdoc" size="20" value="<%=nowrefstockdoc%>">
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
  sql="select * from stockdocument where stockcategory='B'" <!--ֻȡ���ⵥ
  if nowkeyword<>"" then
  	sql=sql&" and stocknumber='"&nowkeyword&"'"
  end if  

  if nowrefstockdoc<>"" then
  	sql=sql&" and refstockdoc='"&nowrefstockdoc&"'"
  end if    
  
  if nowfrom<>"" and nowto="" then
	sql=sql&" and stockdate=#"&nowfrom&"#"
  elseif nowfrom="" and nowto<>"" then
	sql=sql&" and stockdate=#"&nowto&"#"
  elseif nowfrom<>"" and nowto<>"" then
	sql=sql&" and stockdate>=#"&nowfrom&"# and stockdate<=#"&nowto&"#"
  end if

	
  if request("order1")<>"" then
    sql=sql&" order by stocknumber "&request("order1")
  elseif request("order2")<>"" then
    sql=sql&" order by stockstatus "&request("order2")
  elseif request("order3")<>"" then
    sql=sql&" order by refshipment "&request("order3") 
  elseif request("order4")<>"" then
    sql=sql&" order by refitem "&request("order4") 
  elseif request("order5")<>"" then
    sql=sql&" order by contract "&request("order5") 
  elseif request("order6")<>"" then
    sql=sql&" order by material "&request("order6")
  elseif request("order7")<>"" then
    sql=sql&" order by country "&request("order7")
  elseif request("order8")<>"" then
    sql=sql&" order by plant "&request("order8")
  elseif request("order9")<>"" then
    sql=sql&" order by stockdate "&request("order9") 
  elseif request("order10")<>"" then
    sql=sql&" order by coldstorage "&request("order10") 
  elseif request("order11")<>"" then
    sql=sql&" order by productdate "&request("order11")
  elseif request("order12")<>"" then
    sql=sql&" order by customer "&request("order12")	
  end if
  
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;���ⵥ��ѯ</td>
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
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
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
  <tr align="center">
	<td class="category" width="60" height="30">
		<a href="?order1=<%if request("order1")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">���ⵥ��<%if request("order1")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order2=<%if request("order2")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">״̬<%if request("order2")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order2=<%if request("order3")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�ο���ⵥ<%if request("order3")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order9=<%if request("order9")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��������<%if request("order9")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>	
	<td class="category" width="80" height="30">
		<a href="?order11=<%if request("order11")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��������<%if request("order11")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>  
	<td class="category" width="100" height="30">
	  <a href="?order3=<%if request("order4")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�ο�����<%if request("order4")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="100" height="30">
		<a href="?order4=<%if request("order5")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�ο���Ŀ<%if request("order5")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="120" height="30">
		<a href="?order5=<%if request("order6")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">��ͬ��<%if request("order6")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order6=<%if request("order7")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">Ʒ��<%if request("order7")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order8=<%if request("order8")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">����<%if request("order8")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
	</td>
	<td class="category" width="80" height="30">
		<a href="?order10=<%if request("order10")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">�������<%if request("order10")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>
	<td class="category" width="80" height="30">
		<a href="?order12=<%if request("order12")="asc" then%>desc<%else%>asc<%end if%>&form=<%=request("form")%>" class="title">����<%if request("order12")="asc" then%><img src="../images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="../images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>	
  </td>  
  <td class="category">�޸�</td>
  </tr>
  <%
  set rs_stock =server.createobject("ADODB.RecordSet")	
  rs_stock.open sql,conn,1,3
  if not rs_stock.eof then
  do while rs_stock.eof=false
  %>
  <tr align="center">
	  <td align="center" height="30"><%=rs_stock("stocknumber")%></td>	  
	  <td align="center">
		<%
			status_sql="select * from stockstatus where category='B' and status='"&rs_stock("stockstatus")&"'"
			set rs_status=conn.execute(status_sql)
		%>			
	  <%=rs_status("description")%></td>	  
	  <td align="center"><%=rs_stock("refstockdoc")%></td>
	  <td align="center"><%=rs_stock("stockdate")%></td>	  
	  <td align="center"><%=rs_stock("stockqty")%></td>	
	<%
	  if rs_stock("refstockdoc")<>"" then
		  refSql="select * from stockdocument where stockcategory='A' and stocknumber='"&rs_stock("refstockdoc")&"'"
		  set rs_refstock=conn.execute(refSql)	
		  nowrefshipment=rs_refstock("refshipment")
		  nowrefitem=rs_refstock("refitem")
		  nowcontract=rs_refstock("contract")
		  nowmaterial=rs_refstock("material")
		  nowplant=rs_refstock("plant")
		  nowcoldstorage=rs_refstock("coldstorage")
		  nowcustomer=rs_refstock("customer")
	  end if
	%>	  
	  <td align="center"><%=nowrefshipment%></td>
	  <td align="center"><%=nowrefitem%></td>
	  <td align="center"><%=nowcontract%></td>
	  <td align="center"><%=nowmaterial%></td>
	  <td align="center"><%=nowplant%></td>
	  <td align="center"><%=nowcoldstorage%></td>
	  <td align="center"><%=nowcustomer%></td>
	  <td align="center">
			<a href="change_stock_issue.asp?form=<%=request("form")%>&stocknumber=<%=rs_stock("stocknumber")%>&keyword=<%=nowkeyword%>"><img src="../images/res.gif" border="0" hspace="2" align="absmiddle">�޸�</a>
	  </td>
  </tr>
  </tr>
  <%
  	'end if
    rs_stock.movenext
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
  if rs_stock.recordcount>0 then 
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