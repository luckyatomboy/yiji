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
<title><%=dianming%> - �鿴��Ӧ��</title>
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
if fla35="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>

<%
sql="select * from vendor where vendorname='"&request("vendorname")&"'"
set rs=conn.execute(sql)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�鿴��Ӧ������</td>
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
  <form name="form2">
  		<input type="hidden" name="keyword" value="<%=request("keyword")%>">
  		<input type="hidden" name="vendorname" value="<%=request("vendorname")%>">
      <tr>
        <td width="25%" height="30" align="right">��Ӧ�����ƣ�</td>
        <td width="75%" class="category"><%=rs("vendorname")%>
        	</td>
      </tr>    
      <tr>
        <td align="right" height="30">����</td>
        <td class="category">
          <%
          sql="select * from country"
          set rs_country=conn.execute(sql)
          %>
          <select name="country" readonly>
            <%
              do while rs_country.eof=false
            %>
              <option value="<%=rs_country("country")%>" <%if rs_country("country")=rs("country") then%>selected="selected"<%end if%>><%=rs_country("country")%></option>
            <%
              rs_country.movenext
              loop
            %>
          </select>
        </td>
      </tr>        
      <tr>
        <td align="right" height="30">��˾��ַ��</td>
        <td class="category"><input type="text" name="address" style="width:300px;border:none" readonly value="<%=rs("address")%>"></td>
      </tr> 	  
      <tr>
        <td align="right" height="30">��ϵ�绰��</td>
        <td class="category"><input type="text" name="tel" style="width:200px;border:none" readonly value="<%=rs("tel")%>"></td>
      </tr>	 
      <tr>
        <td align="right" height="30">���棺</td>
        <td class="category"><input type="text" name="fax" style="width:200px;border:none" readonly value="<%=rs("fax")%>"></td>
      </tr>	       
	    <tr>
        <td align="right" height="30">Email��</td>
        <td class="category"><input type="text" name="email" style="width:200px;border:none" readonly value="<%=rs("email")%>"></td>
      </tr>    
      <tr>
        <td align="right" height="30">���ʽ1��</td>
        <td class="category"><input type="text" name="term1" style="width:100px;border:none" readonly value="<%=rs("term1")%>"></td>
      </tr> 	
      <tr>
        <td align="right" height="30">���ʽ2��</td>
        <td class="category"><input type="text" name="term2" style="width:100px;border:none" readonly value="<%=rs("term2")%>"></td>
      </tr> 	    
<!--added on 2017/2/4  begin-->
      <tr>
        <td align="right" height="30">����1��</td>
        <td class="category"><input type="text" name="plant1" style="width:60px;border:none" readonly value="<%=rs("plant1")%>"> 
        	&nbsp;&nbsp;&nbsp;����2 <input type="text" name="plant2" style="width:60px;border:none" readonly value="<%=rs("plant2")%>"> 
        	&nbsp;&nbsp;&nbsp;����3 <input type="text" name="plant3" style="width:60px;border:none" readonly value="<%=rs("plant3")%>"> 
        	&nbsp;&nbsp;&nbsp;����4 <input type="text" name="plant4" style="width:60px;border:none" readonly value="<%=rs("plant4")%>"> 
        </td>
      </tr> 	      
<!--end -->
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4" readonly><%=rs("memo")%></textarea></td>
      </tr>	   
      <tr>
        <td align="right" height="30">����ʱ�䣺</td>
        <td class="category"><%=rs("createdate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category"><%=rs("creator")%>
      </tr>	   
      <tr>
        <td align="right" height="30">�޸�ʱ�䣺</td>
        <td class="category"><%=rs("changedate")%>
      </tr>	  
      <tr>
        <td align="right" height="30">�޸��ˣ�</td>
        <td class="category"><%=rs("changer")%>
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