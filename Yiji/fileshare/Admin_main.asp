<%@LANGUAGE="VBScript" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<html>

<head>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<%
Dim Msg
	If Request.QueryString("Action") = "Save" Then SaveData
	Sub SaveData()
	myConn.execute("update fConfig set OKAr='"&Request.Form("ftype")&"',OKsize="&Request.Form("fsize"))
	Msg = "�ɹ��޸����ļ�������Ϣ"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='0;URL=Admin_Main.asp'>"&Msg&"<p align=center>>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_Main.asp>����˴�����</a>")
Response.End()
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ZH������</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
<style>
<!--
#glowtext{
filter:glow(color=Black,strength=2);
width:100%;
}
-->
</style>
<script language="JavaScript1.2">
function glowit(which){
if (document.all.glowtext[which].filters[0].strength==2)
document.all.glowtext[which].filters[0].strength=1
else
document.all.glowtext[which].filters[0].strength=2
}
function glowit2(which){
if (document.all.glowtext.filters[0].strength==2)
document.all.glowtext.filters[0].strength=1
else
document.all.glowtext.filters[0].strength=2
}
function startglowing(){
if (document.all.glowtext&&glowtext.length){
for (i=0;i<glowtext.length;i++)
eval('setInterval("glowit('+i+')",150)')
}
else if (glowtext)
setInterval("glowit2(0)",150)
}
if (document.all)
window.onload=startglowing
</script>
</head>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�ļ�����</td>
	  <td align="right">��</td>
    
    </tr>
  </table>
  </table>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF"  topmargin="0" leftmargin="0" " >

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
<form name="Edit" action="Admin_Main.asp?Action=Save" method="post">
  <tr> 
    <td height="25" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from fConfig"
	frst.open sql,myconn,1,1
	If not frst.Eof then
	%>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr align="center" class="text"> 
            <td height="1" colspan="2" bgcolor="#FFFFFF"><span id="glowtext"></span></td>
          </tr>
          <tr align="center" class="text"> 
            <td height="1" colspan="2" bgcolor="#FFFFFF">[ <a href="index.asp">������ҳ</a>    
              ] [<a href="Admin_List.asp"> �ļ�����</a> ]</td>   
          </tr>
          <tr class="text"> 
            <td width="200"> <div align="right">�����ϴ����ļ����ͣ�</div></td>
            <td><input name="ftype" type="text" class="TextBoxT" value="<%=rs(0)%>" size="50" readonly>
              ��,�ָ���׺��</td>    
          </tr>
          <tr class="text"> 
            <td width="200"><div align="right">�����ϴ����ļ���С��</div></td>
            <td><input name="fsize" type="text" class="TextBoxT" value="<%=rs(1)%>" size="15">
              ��λ:Byte</td>    
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#FFFFFF">��</td>
            <td height="1" bgcolor="#FFFFFF"> <input type="submit" name="Submit" value="�޸�"></td>
          </tr>
        </table>
      <%
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>û�ж�Ӧ�����ݣ�</td>
        </tr>
      </table>
      <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
	  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
            <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>��</td>
	  <td align="right">��</td>
    
    </tr>
  </table>
  </table>

</body>
</html>