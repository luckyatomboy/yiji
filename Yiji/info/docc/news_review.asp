<%@language="VBscript"%>
<%Response.Expires=0%>
<%
id=request.querystring("id")
oabusyusername=request.cookies("oabusyusername")
%>
<title>��������</title>
<style>
	td{font-size:13px; color: #215DC6; }
</style>
<script language="javascript">
	function doSubmit(){
		var obj=document.frmrecommend;

		if(obj.content.value==""){
			alert("�������������ݣ�");
			obj.content.focus();
			return false;
		}
	}
</script>
<body topmargin=0 marginheight=0 leftmargin=0 marginwidth=0>
<br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="24" background="images/bg1.gif"><font style="font-size:13px;" color="#215DC6"><b>����������</b></font></td>
  </tr>
</table>
<table align="center" border="0" cellpadding="1" cellspacing="1" width="90%" bgcolor="e8e8e8">
  <form name="frmrecommend" action="news_review_ok.asp" method="post" onsubmit="javascript:return doSubmit();">
  <input type="hidden" name="action" value="news_review">
  <input type="hidden" name="id" value="<%=request.querystring("id")%>">
    <tr bgcolor="#FFFFFF"> 
      <td height="30" width="120">�������ݣ�</td>
      <td><textarea name="content" style="width:380px;height:60px;"></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30">�������֣�</td>
      <td><input type="text" name="yourname" value="<%=Session("ManageName")%>" style="width:280px;" readonly></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30">�������䣺</td>
      <td><input type="text" name="yourmail" style="width:280px;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" height="26" align="center"> <input type="submit" name="btnSubmit" value="ȷ��" style="width:60px;"> 
        <input type="reset" name="btnReset" value="ȡ��" style="width:60px;"> <input type="button" name="btnClose" onclick="javascript:window.close();" value="�ر�" style="width:60px;"> 
      </td>
    </tr>
  </form>
</table>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="24" background="images/bg1.gif"><font style="font-size:13px;" color="#215DC6"><b>��ע��</b></font></td>
  </tr>
</table>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td height="109" bgcolor="#FFFFFF"><font color=#7B7B7B>�� �������ϵ��£����ء�ȫ���˴�ί�����ά����������ȫ�ľ��������л����񹲺͹����������йط��ɷ���<br>
      �� �������ϵ��£������л����񹲺͹��ĸ����йط��ɷ���<br>
      �� �е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·�������<br>
      �� �������԰������Ա��Ȩ������ɾ�����Ͻ�����е��������� <br>
      �� �������԰巢�����Ʒ����Ȩ����վ��ת�ػ�����<br>
      �� ���뱾���Լ��������Ѿ��Ķ��������������� </font></td>
  </tr>
</table>
