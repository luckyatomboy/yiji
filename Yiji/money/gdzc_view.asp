<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file ="Inc/Date.asp"-->
<%
Dim payid
Dim rs,ID,ListID,Payer,PayType,Money,Project,Menu,Times,InOut
payid = Request("id")
Set rs=Conn.Execute("Select * From [gdzcgl] Where ID="&payid)
ID = rs("ID")
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
<body text="#000000">
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
  <form action="gdzc_Edit2.asp?id=<%=ID%>&action=add" method=post id=form1 name=form1 onSubmit="return chk(this)">
    <tr bgcolor=ffffff> 
      <th height=25 colspan=8 align="center">�̶��ʲ���Ϣ��ʾ</th>
    </tr>
    <tr bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">���</div></td>
      <td width="7%" class="category"> <input name="class" type="text" id="class" value="<%=rs("class")%>" size="10" maxlength="10" readonly></td>
      <td width="8%" class="category"><div align="right">�豸�ţ�</div></td>
      <td width="7%" class="category"> <input name="sb_id" type="text" id="sb_id" value="<%=rs("sb_id")%>" size="10" maxlength="10"  readonly></td>
      <td width="10%" class="category"><div align="right">�ɹ����ڣ�</div></td>
      <td colspan="2" class="category"><input name="caigoudate" type="text" id="caigoudate" value="<%=rs("caigoudate")%>" size="12" maxlength="12" readonly> 
        <input onclick="popUpCalendar(this, form1.caigoudate, 'yyyy-mm-dd')" type="button" value="��ѡ������"></td>
    </tr>
    <tr  bgcolor=ffffff> 

      <td width="8%" class="category"><div align="right">�豸���ƣ�</div></td>
      <td width="7%" class="category"> <input name="name" type="text" id="name" value="<%=rs("name")%>" size="10" maxlength="10"></td>

      <td width="8%" class="category"><div align="right">ʹ�ò��ţ�</div></td>
      <td width="10%" class="category"> <input name="department" type="text" id="department" value="<%=rs("department")%>" size="20" maxlength="150"></td>

      <td width="8%" class="category"><div align="right">������</div></td>
      <td width="7%" class="category"> <input name="quantity" type="text" id="quantity" onkeypress="javascript:CheckNum();" value="<%=rs("quantity")%>" size="10" maxlength="10"></td>


    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">��</div></td>
      <td width="7%" class="category"> <input name="money" type="text" id="money" value="<%=rs("money")%>" size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right">�۾��ʣ�</div></td>
      <td width="7%" class="category"> <input name="zejiu" type="text" id="zejiu"  onkeypress="javascript:CheckNum();" value="<%=rs("zejiu")%>" size="10" maxlength="10">%</td>

      <td  width="8%" class="category"><div align="right">��ע��</div></td>
      <td width="17%" class="category"> <input name="remark" type="text" id="remark" value="<%=rs("remark")%>" size="20" maxlength="50"></td>



    </tr>

    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">״̬��</div></td>
      <td width="7%" class="category"> <input name="score" type="text" id="score" value="<%=rs("score")%>" size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>

      <td  width="8%" class="category"><div align="right"></div></td>
      <td width="17%" class="category"></td>



    </tr>

  </form>
</table>
<%
rs.Close
Set rs=Nothing

End if%>
</html>