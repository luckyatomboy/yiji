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
<title><%=dianming%> - ������Ϣ����</title>
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
if fla59="0" and session("redboy_id")<>"1" then
%>
<br><center><img src="../images/note.gif" align="absmiddle">&nbsp;<font color="#FF0000">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>
<%  
  response.end
end if
%>
<%if request("hid1")="" then%>
<form name="form1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;������Ϣ����</td>
	  <td align="right">��</td>
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
    <td width="28%" height="30" align="right"><b>��˾���ƣ�</b> </td>
    <td width="72%" class="category">
        <input type="text" name="dianming" style="width:300px" value="<%=dianming%>"></td>
  </tr>
   <tr>
    <td width="28%" height="30" align="right"><b>ϵͳ���棺</b> </td>
    <td width="72%" class="category">
        <input type="text" name="gonggao" style="width:300px" value="<%=gonggao%>"></td>
  </tr>

  <tr>
    <td width="28%" height="30" align="right"><b>����һ����Ա�Ӷ��ٻ��֣�</b> </td>
    <td width="72%" class="category">
        <input type="text" name="jieshaojifen" style="width:200px" value="<%=jieshaojifen%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')">
        <font color="#666666">(ֻ���ǰ���������)</font> </td>
  </tr>
  <tr>
    <td width="28%" height="30" align="right"><b>����Ӷ��ٻ��֣�</b> </td>
    <td width="72%" class="category">
        <input type="text" name="xuhuijifen" style="width:200px" value="<%=xuhuijifen%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')">
        <font color="#666666">(ֻ���ǰ���������)</font></td>
  </tr> 
  <tr>
    <td width="28%" height="30" align="right"><b>�Ƿ����ÿ�汨�����ܣ�</b> </td>
    <td width="72%" class="category">
        <input type="checkbox" name="baojin1"<%if baojin="yes" then%> checked="checked"<%end if%> value="yes"></td>
  </tr> 
  <tr>
    <td width="28%" height="30" align="right"><b>�Ƿ����û�Ա�������ѹ��ܣ�</b> </td>
    <td width="72%" class="category">
        <input type="checkbox" name="baojin2"<%if baojin2="yes" then%> checked="checked"<%end if%> value="yes"></td>
  </tr> 
  <tr>
    <td width="28%" height="30" align="right"><b>��Ա������ǰ���������ѣ�</b> </td>
    <td width="72%" class="category">
        <input type="text" name="tiqian" style="width:200px" value="<%=tiqian%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')">
        <font color="#666666">(ֻ���ǰ���������)</font></td>
  </tr>
  <tr class="a3">
    <td width="28%" height="30" align="right"><b>��Ӳ�Ʒ���Ƿ����ϴ�ӡ��</b> </td>
    <td width="72%" class="category">
        <input type="checkbox" name="dayin1"<%if dayin1="yes" then%> checked="checked"<%end if%> value="yes"></td>
  </tr> 
  <tr>
    <td width="28%" height="30" align="right"><b>���۲�Ʒ���Ƿ����ϴ�ӡ��</b> </td>
    <td width="72%" class="category">
        <input type="checkbox" name="dayin2"<%if dayin2="yes" then%> checked="checked"<%end if%> value="yes"></td>
  </tr>
  <tr>
    <td width="28%" height="30" align="right"><b>���ֲ�ѯ��ÿҳ��ʾ��������¼��</b> </td>
    <td width="72%" class="category">
        <input type="text" name="maxrecord" style="width:200px" value="<%=maxrecord%>" onKeyUp="this.value=this.value.replace(/[^\d.]/g,'')"  onafterpaste="this.value=this.value.replace(/[^\d.]/g,'')">
        <font color="#666666">(ֻ���ǰ���������)</font></td>
  </tr>    
  <tr>
    <td height="30">��</td>
    <td class="category">
        <input name="submit" type="submit" value=" ȷ �� " class="button">
      &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="hidden" name="hid1" value="ok">
        <input name="reset" type="reset" value=" ������д " class="button">
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
nowdianming=request("dianming")
nowgonggao=request("gonggao")
nowjieshaojifen=request("jieshaojifen")
if nowjieshaojifen="" then
  nowjieshaojifen=0
end if
nowxuhuijifen=request("xuhuijifen")
if nowxuhuijifen="" then
  nowxuhuijifen=0
end if
maxrecord=request("maxrecord")
if maxrecord="" then
  maxrecord=16
end if
maxproduit=request("maxproduit")
if maxproduit="" then
  maxproduit=1
end if
nowbaojin=request("baojin")
if nowbaojin="" then
  nowbaojin="no"
end if
nowbaojin2=request("baojin2")
if nowbaojin2="" then
  nowbaojin2="no"
end if
tiqian=request("tiqian")
if tiqian="" then
  tiqian=2
end if
nowdayin1=request("dayin1")
if nowdayin1="" then
  nowdayin1="no"
end if
nowdayin2=request("dayin2")
if nowdayin2="" then
  nowdayin2="no"
end if
nowshowpic=request("showpic")
if nowshowpic="" then
  nowshowpic="no"
end if
sql="update config set dianming='"&nowdianming&"',gonggao='"&nowgonggao&"',jieshaojifen="&nowjieshaojifen&",xuhuijifen="&nowxuhuijifen&",baojin='"&nowbaojin&"',baojin2='"&nowbaojin2&"',tiqian="&tiqian&",dayin1='"&nowdayin1&"',dayin2='"&nowdayin2&"',maxrecord="&maxrecord&" where id=1"
conn.execute(sql)
Function finddir(filepath)
	finddir=left(filepath,instrRev(filepath,"/")-7)
end Function
theurl="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
%>
<script language="javascript">
alert("������Ϣ���óɹ���")
window.location.href="config.asp"
</script> 
<%
end if
%>
</body>
</html>