<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file ="Inc/Date.asp"-->
<%
Dim isInOut

isInOut = Request("InOut")




%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script language="JavaScript" src="Image/js.js"></SCRIPT>
</head>
<%
Dim payid
Dim rs,ID,ListID,Payer,PayType,Money,Project,Menu,Times,InOut
payid = Request("id")
Set rs=Conn.Execute("Select * From [gdzcgl] Where 1=2")
%>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="Image/style.css" type=text/css rel=stylesheet>
<script language="JavaScript" src="Image/js.js"></SCRIPT>
</head>
<%
If Request.QueryString("action")="update" Then
   Call PaySave
Else
%>

<script language="JavaScript">
<!--
function change(obj,i) {
he=parseInt(obj.style.height);
if (he>=80&&he<=400)
   obj.style.height=he+i+'px';
else 
   obj.style.height='80px';
}

function chk(theForm){

if (theForm.sbclass.value == ""){
                alert("�������豸���!");
                theForm.sbclass.focus();
                return (false);
        }         

if (theForm.sb_id.value == ""){
                alert("�������豸���!");
                theForm.sb_id.focus();
                return (false);
        }         

if (theForm.name.value == ""){
                alert("�������豸����!");
                theForm.name.focus();
                return (false);
        }         

if (theForm.department.value == ""){
                alert("�����벿��!");
                theForm.department.focus();
                return (false);
        } 
if (theForm.money.value == ""){
                alert("��������!");
                theForm.money.focus();
                return (false);
        } 
if (theForm.quantity.value == ""){
                alert("����������!");
                theForm.quantity.focus();
                return (false);
        } 
if (theForm.zejiu.value == ""){
                alert("�������۾�!");
                theForm.zejiu.focus();
                return (false);
        } 

if (theForm.score.value == ""){
                alert("������״̬!");
                theForm.score.focus();
                return (false);
        } 
if (theForm.zejiu.value*1 > 100){
                alert("�۾ɲ��ܳ���100%!");
                theForm.zejiu.focus();
                return (false);
        } 

if (theForm.score.value!="����" && theForm.score.value!="����"){
                alert("״ֻ̬��Ϊ�����򱨷�!");
                theForm.score.focus();
                return (false);
        } 
}

function CheckNum(){
   if((event.keyCode<48||event.keyCode>57)&&event.keyCode!=46) event.returnValue=false; 
   }
//-->
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;�̶��ʲ���Ϣ����</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
<body text="#000000">
<table cellpadding="2" cellspacing="1" class="toptable grid" border="1" width="95%" class="tableBorder" align=center>
  <form action="gdzc_add2.asp?id=<%=ID%>&action=add" method=post id=form1 name=form1 onSubmit="return chk(this)">

    <tr bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">���</div></td>
      <td width="7%" class="category"> <input name="sbclass" type="text" id="sbclass"  size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right">�豸�ţ�</div></td>
      <td width="7%" class="category"> <input name="sb_id" type="text" id="sb_id"  size="10" maxlength="10"></td>
      <td width="10%" class="category"><div align="right">�ɹ����ڣ�</div></td>
      <td colspan="2" class="category"><input name="caigoudate" type="text" id="caigoudate" value="<%=date()%>" size="12" maxlength="12" readonly> 
        <input onclick="popUpCalendar(this, form1.caigoudate, 'yyyy-mm-dd')" type="button" value="��ѡ������"></td>
    </tr>
    <tr  bgcolor=ffffff> 

      <td width="8%" class="category"><div align="right">�豸���ƣ�</div></td>
      <td width="7%" class="category"> <input name="name" type="text" id="name"  size="10" maxlength="10"></td>

      <td width="8%" class="category"><div align="right">ʹ�ò��ţ�</div></td>
      <td width="10%" class="category"> <input name="department" type="text" id="department"  size="20" maxlength="150"></td>

      <td width="8%" class="category"><div align="right">������</div></td>
      <td width="7%" class="category"> <input name="quantity" type="text" id="quantity" onkeypress="javascript:CheckNum();"  size="10" maxlength="10"></td>


    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">��</div></td>
      <td width="7%" class="category"> <input name="money" type="text" id="money" onkeypress="javascript:CheckNum();"  size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right">�۾��ʣ�</div></td>
      <td width="7%" class="category"> <input name="zejiu" type="text" id="zejiu"  onkeypress="javascript:CheckNum();" value="0" size="10" maxlength="10">%</td>

      <td  width="8%" class="category"><div align="right">��ע��</div></td>
      <td width="17%" class="category"> <input name="remark" type="text" id="remark"  size="20" maxlength="50"></td>



    </tr>

    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">״̬��</div></td>
      <td width="7%" class="category"> <input name="score" type="text" id="score" value="����" size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>

      <td  width="8%" class="category"><div align="right"></div></td>
      <td width="17%" class="category"></td>



    </tr>

    <tr  bgcolor=ffffff> 
      <td colspan="7" align="center" class="category"><input type="submit" name="Submit" value="�� ��"></td>
    </tr>
  </form>
</table>
<%
rs.Close
Set rs=Nothing

End if%>
</html>