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
Set rs=Conn.Execute("Select * From [XJLB] Where ID="&payid)
ID = rs("ID")
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
if (theForm.XJCODE.value == ""){
                alert("�������Ŀ����!");
                theForm.XJCODE.focus();
                return (false);
        } 

if (theForm.XJNAME.value == ""){
                alert("�������Ŀ����!");
                theForm.XJNAME.focus();
                return (false);
        } 
if (theForm.SCORE.value == ""){
                alert("���������!");
                theForm.SCORE.focus();
                return (false);
        }         
if (theForm.qc.value == ""){
                alert("������������!");
                theForm.qc.focus();
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
      <td>&nbsp;��������޸Ŀ�Ŀ</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
<body text="#000000">
<table cellpadding="2" cellspacing="1" class="toptable grid" border="1" width="95%"  align=center>
  <form action="cwlb_edit2.asp?id=<%=ID%>&action=add" method=post id=form1 name=form1 onSubmit="return chk(this)">

    <tr bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">��־��</div></td>
      <td width="12%" class="category"> 
      	<select name="SCORE" id="SCORE">
          <%
   Response.write "<option value='��'>��</option>"
   Response.write "<option value='��'>��</option>"

%>
        </select>
����:<%=rs("Score")%>      	
      	</td>
      <td width="8%" class="category"><div align="right">��Ŀ���룺</div></td>
      <td width="7%" class="category"> <input name="XID" type="hidden" id="XID" value="<%=rs("ID")%>""><input name="XJCODE" type="text" id="XJCODE" value="<%=rs("XJCODE")%>" size="10" maxlength="10" onKeyPress="javascript:CheckNum();"></td>
      <td width="10%" class="category"><div align="right">��Ŀ���ƣ�</div></td>
      <td colspan="2" class="category"><input name="XJNAME" type="text" id="XJNAME" value="<%=rs("XJNAME")%>" size="12" maxlength="12"> 

    </tr>
    <tr  bgcolor=ffffff> 

      <td width="8%" class="category"><div align="right">�������ƣ�</div></td>
      <td width="7%" class="category"> <input name="XJYH" type="text" id="XJYH" value="<%=rs("XJYH")%>" size="10" maxlength="10"></td>

      <td width="8%" class="category"><div align="right">�����˺ţ�</div></td>
      <td width="10%" class="category"> <input name="YHZH" type="text" id="YHZH" value="<%=rs("YHZH")%>" size="20" maxlength="150"></td>

      <td width="8%" class="category"><div align="right">��ϵ�ˣ�</div></td>
      <td width="7%" class="category"> <input name="YHLXR" type="text" id="YHLXR" value="<%=rs("YHLXR")%>" size="10" maxlength="10"></td>


    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">��ϵ�绰��</div></td>
      <td width="7%" class="category"> <input name="XGDH" type="text" id="XGDH" value="<%=rs("XGDH")%>" size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right">���</div></td>
      <td width="7%" class="category"> <input name="LB" type="text" id="LB" value="<%=rs("LB")%>" size="10" maxlength="10"></td>

      <td  width="8%" class="category"><div align="right">��ע��</div></td>
      <td width="17%" class="category"> <input name="remark" type="text" id="remark" value="<%=rs("remark")%>" size="20" maxlength="50"></td>
    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">�ڳ���</div></td>
      <td width="7%" class="category"> <input name="qc" type="text" id="qc" value="<%=rs("qc")%>" size="10" maxlength="10" onKeyPress="javascript:CheckNum();"></td>
      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>

      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>

    <tr  bgcolor=ffffff> 
      <td colspan="7" align="center" class="category"><input type="submit" name="Submit" value="�޸�"></td>
    </tr>
  </form>
</table>
<%
rs.Close
Set rs=Nothing

End if%>
</html>