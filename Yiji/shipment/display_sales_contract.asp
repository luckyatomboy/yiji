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
<title><%=dianming%> - �鿴������ͬ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="../js/jquery-3.2.1.js"></script>
<script type="text/javascript" src="../js/jquery-ui.js"></script>
<link href="../style/jquery-ui.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%
if fla27="0" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if


sql="select * from SalesContract where ContractNum="&request("ContractNum")
set rs=conn.execute(sql)
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;������ͬ&nbsp;<%=rs("ContractNum")%> </td>
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
        <td align="right" height="30">��ͬ���ͣ�</td>
        <td class="category">
			<select name="category" readonly>
				<option value="A" <%if rs("status")="A" then%>selected="selected"<%end if%>>�ֻ�</option>
				<option value="B" <%if rs("status")="B" then%>selected="selected"<%end if%>>�ڻ�</option>
			</select>				
	    </td>
	  </tr>	  
      <tr>
        <td align="right" height="30">�ҷ���˾��</td>
        <td class="category">
			<%
			sql="select * from owncompany order by company"
			set rs_owncompany=conn.execute(sql)
			%>		  		
			<select name="owncompany" readonly>			
				<%
					do while rs_owncompany.eof=false
				%>
					<option value="<%=rs_owncompany("company")%>" <%if rs("owncompany")=rs_owncompany("company") then%>selected="selected"<%end if%>><%=rs_owncompany("company")%></option>
				<%
					rs_owncompany.movenext
					loop
					rs_owncompany.close
				%>								
			</select>			
			&nbsp;&nbsp;&nbsp;&nbsp;�ͻ�
			<%
			sql="select * from customer order by customername"
			set rs_customer=conn.execute(sql)
			%>		  		
			<select name="customer" readonly>			
				<%
					do while rs_customer.eof=false
				%>
					<option value="<%=rs_customer("customername")%>" <%if rs("customer")=rs_customer("customername") then%>selected="selected"<%end if%>><%=rs_customer("customername")%></option>
				<%
					rs_customer.movenext
					loop
					rs_customer.close
				%>								
			</select>	
	    </td>    		
	  </tr>	  
      <tr>
        <td width="20%" align="right" height="30">״̬��</td>
        <td width="80%" class="category">
			<%
			sql="select * from contractstatus"
			set rs_status=conn.execute(sql)
			%>
			<select name="status" readonly>
				<%do while rs_status.eof=false%>
				<option value="<%=rs_status("status")%>" <%if rs("status")=rs_status("status") then%>selected="selected"<%end if%>><%=rs_status("description")%></option>
				<%
					rs_status.movenext
					loop
				%>
			</select>				
	    </td>
	  </tr>
      <tr>	  
	    	<td align="right" height="30">�ο����ڱ�</td>
        <td class="category">
					<input name="refshipment" readonly style="width:100px;border:none" value="<%=rs("refshipment")%>"> 
					&nbsp;&nbsp;&nbsp;��Ŀ��	
					<input name="refitem" readonly style="width:50px;border:none" value="<%=rs("refitem")%>"> 
				</td>
      </tr>   
      <tr>	  
	    <td align="right" height="30">���ţ�</td>
        <td class="category">
			<input name="plant" style="width:100px;border:none" readonly value="<%=rs("plant")%>">
			&nbsp;&nbsp;&nbsp;����
			<input name="guobie" style="width:100px;border:none" readonly value="<%=rs("country")%>">
		</td>
      </tr>       
      <tr>	  
	    <td align="right" height="30">Ʒ����</td>
		<td class="category">
          		<input name="material" style="width:100px;border:none" readonly value="<%=rs("material")%>">
		  		&nbsp;&nbsp;&nbsp;���
		  		<input name="spec" style="width:100px;border:none" readonly value="<%=rs("spec")%>">
		  		&nbsp;&nbsp;&nbsp;��װ
		  		<input name="package" style="width:200px;border:none" readonly value="<%=rs("package")%>">		  		
		</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">������</td>
        <td class="category">
		  		<input name="quantity" style="width:60px;border:none" readonly maxlength="5" value="<%=rs("quantity")%>" >
		  		&nbsp;&nbsp;&nbsp;����
		  		<input name="weight" style="width:50px;border:none" readonly maxlength="20" value="<%=rs("weight")%>">��
		  		&nbsp;&nbsp;&nbsp;����
		  		<input name="price" style="width:50px;border:none" readonly value="<%=rs("price")%>">Ԫ/��		  		
		</td>
      </tr>	  
      <tr>	  
	    <td align="right" height="30">�����⣺</td>
        <td class="category">
		  		<input name="coldstorage" style="width:100px;border:none" readonly value="<%=rs("coldstorage")%>">
				&nbsp;&nbsp;&nbsp;������
		  		<input name="deliveryloc" style="width:100px;border:none" readonly value="<%=rs("deliveryloc")%>">
				</td>
      </tr>
      <tr>	  
	    <td align="right" height="30">Ԥ�Ƶ����ڣ�</td>
        <td class="category">
		  		<input name="boarddate" style="width:80px;border:none" readonly value="<%=rs("boarddate")%>" id="boarddate">	  		
				&nbsp;&nbsp;&nbsp;������
		  		<input name="deliveryport" style="width:100px;border:none" readonly value="<%=rs("deliveryport")%>">							
				</td>
      </tr>
      <tr>
            <td height="30">&nbsp;</td>
              <td class="category">
            <input type="button" value=" ���� " onClick="window.history.go(-1);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;        
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

</body>
</html>