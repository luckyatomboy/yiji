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
<title><%=dianming%> - ���ⵥ</title>
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
<!--if fla1="0" and session("redboy_id")<>"1" then-->
if fla31="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<%
if request("hid1")="ok" then
nowstatus=request("status")
nowrefstockdoc=request("refstockdoc")
nowcategory="B"
if request("issuedate")<>"" then
	nowissuedate=request("issuedate")
else
	nowissuedate=date()
end if
if request("issueqty")<>"" then
	nowissueqty=request("issueqty")
else
	nowissueqty=0
end if


sql="select top 1 * from stockdocument where stockcategory='B' order by stocknumber desc"

set rs_count=conn.execute(sql)

if rs_count.bof and rs_count.eof then
	nowstocknumber="20000001"
else
	rs_count.movefirst
	nowstocknumber=rs_count("stocknumber") + 1
end if	

sql="insert into stockdocument(stocknumber,stockcategory,stockstatus,stockdate,stockqty,refstockdoc,createdate,creator)"
sql=sql&" values("&nowstocknumber&",'"&nowcategory&"','"&nowstatus&"',#"&nowissuedate&"#,"&nowissueqty&",'"&nowrefstockdoc&"',#"&now()&"#,'"&session("redboy_username")&"')"
conn.execute(sql)
%>
<script language="javascript">
alert("���ⵥ¼��ɹ���")
window.location.href="shipment.asp"
</script>

<%
else
%>
<script language="javascript">
function check1(){
<!-- �����ⵥ�Ƿ���� -->
if (document.form1.refstockdoc.value=="����ѡ����ⵥ")
{
alert("��ѡ����ⵥ��");
return false;
}
<!-- �����������Ƿ�Ϊ0 -->
if (document.form1.issueqty.value=="0")
{
alert("���������������");
return false;
}
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;���ⵥ </td>
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
	  <form name="form1">
      <tr>
        <td width="20%" align="right" height="30">״̬��</td>
        <td width="80%" class="category">
					<%
					sql="select * from stockstatus where category='A'"
					set rs_status=conn.execute(sql)
					%>
					<select name="status">
						<%do while rs_status.eof=false%>
						<option value="<%=rs_status("status")%>"><%=rs_status("description")%></option>
						<%
							rs_status.movenext
							loop
						%>
					</select>				
			  </td>
			</tr>
      <tr>	  
	    	<td align="right" height="30">�ο���ⵥ��</td>
        <td class="category">
					<input name="refstockdoc" readonly style="width:100px" value="����ѡ����ⵥ" onClick="JavaScript:window.open('query_stock_order.asp?queryform=form1&field=refstockdoc&field2=refshipment&field3=refitem&field4=stockdate&field5=plant&field6=contract&field7=material&field8=spec&field9=price&field10=currency&field11=case&field12=quantity&field13=weight&field14=coldstorage&field15=customer&field16=stockqty&field17=reason&field18=remainqty&field19=productdate&field20=warrantyperiod','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1200,height=470,top=100,left=20');">
				</td>
      </tr>			
      <tr>	  
	    	<td align="right" height="30">����������</td>
        <td class="category">
		  		<input name="issueqty" style="width:100px" value=0>			
				&nbsp;&nbsp;&nbsp;��������
				<input name="issuedate" readonly style="width:80px">
		  		<img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('../day.asp?form=form1&field=issuedate&oldDate='+issuedate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">				
				</td>
      </tr>	  
      <tr>	  
	    	<td align="right" height="30">�ο����ڱ�</td>
        <td class="category">
					<input name="refshipment" style="width:100px" disabled>
					&nbsp;&nbsp;&nbsp;��Ŀ
					<input name="refitem" style="width:100px" disabled>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">������ڣ�</td>
        <td class="category">
				<input name="stockdate" disabled style="width:80px">
				</td>
      </tr>     
      <tr>	  
	    	<td align="right" height="30">���ţ�</td>
        <td class="category">
					<input name="plant" style="width:100px" disabled>
				&nbsp;&nbsp;&nbsp;��ͬ��	
				<input name="contract" style="width:150px" disabled>
				</td>
      </tr>       
      <tr>	  
	    	<td align="right" height="30">Ʒ����</td>
			<td class="category">
				<input name="material" style="width:100px" disabled>
		  		&nbsp;&nbsp;&nbsp;���
		  		<input name="spec" style="width:100px" disabled>
		  		&nbsp;&nbsp;&nbsp;����
		  		<input name="price" style="width:100px" disabled>&nbsp;/��
		  		&nbsp;&nbsp;&nbsp;����
				<input name="currency" style="width:100px" disabled>
			</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">��ţ�</td>
        <td class="category">
		  		<input name="case" style="width:120px" disabled>
		  		&nbsp;&nbsp;&nbsp;����
		  		<input name="quantity" style="width:80px" disabled>
		  		&nbsp;&nbsp;&nbsp;����
		  		<input name="weight" style="width:150px" disabled>
				</td>
      </tr>	  
      <tr>	  
	    	<td align="right" height="30">������ƣ�</td>
        <td class="category">
		  		<input name="coldstorage" style="width:100px" disabled>
				&nbsp;&nbsp;&nbsp;����
		  		<input name="customer" style="width:100px" disabled>
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">���������</td>
        <td class="category">
		  		<input name="stockqty" style="width:100px" disabled>
				&nbsp;&nbsp;&nbsp;���ԭ��
		  		<input name="reason" style="width:100px" disabled>				
				&nbsp;&nbsp;&nbsp;ʣ������
		  		<input name="remainqty" style="width:100px" disabled >				
				</td>
      </tr>
      <tr>	  
	    	<td align="right" height="30">�������ڣ�</td>
        <td class="category">
		  		<input name="productdate" style="width:80px" disabled>
		  		&nbsp;&nbsp;&nbsp;&nbsp;������
				<input name="warrantyperiod" style="width:100px" disabled>
				</td>
      </tr> 
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" ȷ��¼�� " onClick="return check1()" class="button">
		  <input type="hidden" name="hid1" value="ok">
		  <input type="reset" value=" ������д " class="button">
		  </td>
      </tr>
	  </form>
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
<%end if%>
</body>
</html>