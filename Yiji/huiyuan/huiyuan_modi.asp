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
<title><%=dianming%> - ��ӻ�Ա</title>
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
if fla37="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%if request("hid1")="" then%> 
<script language="javascript">
function check()
{
if (document.form1.username.value==""||document.form1.card.value=="")
{
alert("��*�ŵı�����д��");
return false;
}
}
</script>
<%
sql="select * from huiyuan where id="&request("id")
set rs=conn.execute(sql)
%>
<form name="form1">
  <input type="hidden" name="form" value="<%=request("form")%>">
  <input type="hidden" name="field" value="<%=request("field")%>">
  <input type="hidden" name="field2" value="<%=request("field2")%>">
  <input type="hidden" name="field3" value="<%=request("field3")%>">
  <input type="hidden" name="page" value="<%=request("page")%>">
  <input type="hidden" name="zu" value="<%=request("zu")%>">
  <input type="hidden" name="keyword" value="<%=request("keyword")%>">
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
  <input type="hidden" name="order14" value="<%=request("order14")%>">
  <input type="hidden" name="order15" value="<%=request("order15")%>">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�޸Ļ�Ա����(��*�ŵ�Ϊ������)</td>
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
        <td width="25%" height="30" align="right">��Ա�飺</td>
        <td width="75%" class="category">
	<%
	sql="select * from zu_huiyuan order by id"
	set rs_zu=conn.execute(sql)
	if rs_zu.eof then
	%>
	<script language="javascript">
	  alert("������ӻ�Ա�飡")
	  window.location.href="zu_add.asp"
	</script>
	<%
	response.end
	end if
	%>
	  <select name="id_zu">
        <%
	do while rs_zu.eof=false
	%>
        <option value="<%=rs_zu("id")%>"<%if rs_zu("id")=rs("id_zu") then%> selected="selected"<%end if%>><%=rs_zu("zu")%></option>
        <%
	  rs_zu.movenext
	loop
	%>
      </select>
		</td>
      </tr>
	  <tr>
        <td align="right" height="30">��Ա���ţ�</td>
        <td class="category">
            <input type="text" name="card" style="width:300px" value="<%=rs("card")%>">&nbsp;&nbsp;<font color="#ff0000">*</font></td>
      </tr>
	  <tr>
        <td height="30" align="right">��Ա������</td>
        <td class="category">
            <input type="text" name="username" style="width:200px" value="<%=rs("username")%>">
        &nbsp;<font color="#ff0000">*</font></td>
      </tr>
      <tr>
        <td align="right" height="30">��Ա�Ա�</td>
        <td class="category">
          <input type="radio" name="xinbie" value="Ů"<%if rs("xinbie")="Ů" then%> checked="checked"<%end if%>>Ů <input type="radio" name="xinbie" value="��"<%if rs("xinbie")="��" then%> checked="checked"<%end if%>>��</td>
      </tr>
      <tr>
        <td align="right" height="30">��ϵ�绰��</td>
        <td class="category">
            <input type="text" name="tel" style="width:200px" value="<%=rs("tel")%>"></td>
      </tr>
	  <tr>
        <td align="right" height="30">QQ��</td>
        <td class="category">
            <input type="text" name="qq" style="width:200px" value="<%=rs("qq")%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')"></td>
      </tr>
	  <tr>
        <td align="right" height="30">Email��</td>
        <td class="category">
            <input type="text" name="email" style="width:200px" value="<%=rs("email")%>"></td>
      </tr>
      <tr>
        <td align="right" height="30">��ͥסַ��</td>
        <td class="category">
            <input type="text" name="address" style="width:300px" value="<%=rs("address")%>"></td>
      </tr>
      <tr>
        <td align="right" height="30">���֤�ţ�</td>
        <td class="category">
            <input type="text" name="sfz" style="width:300px" value="<%=rs("sfz")%>"></td>
      </tr>
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category">
        <%
		if rs("jieshao")=0 then
		  nowjieshao="��"
		else
		  sql="select * from huiyuan where id="&rs("jieshao")
		  set rs_jieshao=conn.execute(sql)
		  nowjieshao=rs_jieshao("username")
		end if
		%>
		  <%=nowjieshao%>
		</td>
      </tr>
      <tr>
        <td align="right" height="30">�����ˣ�</td>
        <td class="category">
		  <%
		  if session("redboy_id")=1 then
			sql="select * from login order by id_zu,id"
			set rs_login=conn.execute(sql)
		  else
			sql="select * from login where id="&session("redboy_id")
			set rs_login=conn.execute(sql)	  
		  end if
		  if rs_login.eof then
		  %>
		    <a href="user_add.asp">�����Ա��</a>
		  <%
		  else
		  %>
		  <select name="login">
		  <%
		  do while rs_login.eof=false
		  %>
		    <option value="<%=rs_login("username")%>"<%if rs_login("username")=rs("login") then%> selected="selected"<%end if%>><%=rs_login("username")%> (<%=rs_login("bianhao")%>)</option>
		  <%
		    rs_login.movenext
		  loop
		  %>
		  </select>
		  <%
		  end if
		  %>
        </td>
      </tr>
      <tr>
        <td align="right" height="30">��Ա���գ�</td>
        <td class="category">
		  <input name="shenri" value="<%=rs("shenri")%>" readonly style="width:150px">
		  <img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('day.asp?form=form1&field=shenri&oldDate='+shenri.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=470,left=520');">
		  <select name="yinyan">
		    <option value="��"<%if rs("yinyan")="��" then%> selected="selected"<%end if%>>����</option>
			<option value="��"<%if rs("yinyan")="��" then%> selected="selected"<%end if%>>����</option>
		  </select>	
		</td>
      </tr>		  
      <tr>
        <td align="right" height="30">���ʱ�䣺</td>
        <td class="category">
		  <input name="startdate" value="<%=rs("startdate")%>" readonly style="width:100px">	
		  <img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('day.asp?form=form1&field=startdate&oldDate='+startdate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">          
		</td>
      </tr>
      <tr>
        <td align="right" height="30">����ʱ�䣺</td>
        <td class="category">
		  <input name="enddate" value="<%=rs("enddate")%>" readonly style="width:100px">
		  <img src="../images/date.gif" align="absmiddle" style="cursor:pointer;" onClick="JavaScript:window.open('day.asp?form=form1&field=enddate&oldDate='+enddate.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');">          
		</td>
      </tr>	  
      <tr>
        <td align="right" height="30">��Ա���֣�</td>
        <td class="category">
            <input type="text" name="jifen" style="width:300px" value="<%=rs("jifen")%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')"></td>
      </tr>
      <tr>
        <td align="right" height="30">��ע��</td>
        <td class="category"><textarea name="beizhu" cols="70" rows="4"><%=rs("beizhu")%></textarea></td>
      </tr>	 	  	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		  <input type="submit" value=" ȷ���޸� " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="reset" value=" ���� " class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="button" value=" �����޸ķ��� " onClick="window.history.go(-1)" class="button">
		  <input type="hidden" name="hid1" value="ok">
		  <input type="hidden" name="id" value="<%=request("id")%>">
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
<%
else
nowqq=request("qq")
nowemail=request("email")
nowid_zu=request("id_zu")
nowusername=request("username")
nowxinbie=request("xinbie")
nowtel=request("tel")
nowaddress=request("address")
nowsfz=request("sfz")
nowcard=request("card")
nowlogin=request("login")
nowjifen=request("jifen")
nowbeizhu=request("beizhu")
nowstartdate=request("startdate")
nowenddate=request("enddate")
nowshenri=request("shenri")
nowyinyan=request("yinyan")

sql="select * from huiyuan where card='"&nowcard&"' and id<>"&request("id")
set rs=conn.execute(sql)
if rs.eof=false then
%>
<script language="javascript">
alert("������Ŀ����Ѿ�����,���������룡")
window.history.go(-1)
</script> 
<%
  response.end
end if

sql="update huiyuan set card='"&nowcard&"',username='"&nowusername&"',xinbie='"&nowxinbie&"',tel='"&nowtel&"',address='"&nowaddress&"',sfz='"&nowsfz&"',login='"&nowlogin&"',jifen="&nowjifen&",beizhu='"&nowbeizhu&"',startdate=#"&nowstartdate&"#,enddate=#"&nowenddate&"#,shenri=#"&nowshenri&"#,yinyan='"&nowyinyan&"',qq='"&nowqq&"',email='"&nowemail&"',id_zu="&nowid_zu&" where id="&request("id")
conn.execute(sql)
%>
<script language="javascript">
alert("��Ա�����޸ĳɹ���")
window.location.href="huiyuan.asp?form=<%=request("form")%>&field=<%=request("field")%>&field2=<%=request("field2")%>&field3=<%=request("field3")%>&page=<%=request("page")%>&zu=<%=request("zu")%>&keyword=<%=request("keyword")%>&order1=<%=request("order1")%>&order2=<%=request("order2")%>&order3=<%=request("order3")%>&order4=<%=request("order4")%>&order5=<%=request("order5")%>&order6=<%=request("order6")%>&order7=<%=request("order7")%>&order8=<%=request("order8")%>&order9=<%=request("order9")%>&order10=<%=request("order10")%>&order11=<%=request("order11")%>&order12=<%=request("order12")%>&order13=<%=request("order13")%>&order14=<%=request("order14")%>&order15=<%=request("order15")%>"
</script> 
<%
end if
%>
</body>
</html>