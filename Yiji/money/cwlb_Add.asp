<!--#include file="Conn.asp"-->
<!--#include file="Check.asp"-->
<!--#include file ="Inc/Date.asp"-->
<html>
<head>
<title>管理中心</title>
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
%>
<html>
<head>
<title>管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../style/style.css" rel="stylesheet" type="text/css">
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
if (theForm.XJCODE.value == ""){
                alert("请输入科目代码!");
                theForm.XJCODE.focus();
                return (false);
        } 

if (theForm.XJNAME.value == ""){
                alert("请输入科目名称!");
                theForm.XJNAME.focus();
                return (false);
        } 
if (theForm.SCORE.value == ""){
                alert("请输入类别!");
                theForm.SCORE.focus();
                return (false);
        }         
if (theForm.qc.value == ""){
                alert("请输入其初金额!");
                theForm.qc.focus();
                return (false);
        } 
}

function CheckNum(){
   if((event.keyCode<48||event.keyCode>57)&&event.keyCode!=46) event.returnValue=false; 
   }
//-->
</script>
<body text="#000000">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../images/r_1.gif" alt="" /></td>
<td width="100%" background="../images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;财务管理新增科目</td>
	  <td align="right">&nbsp;</td>
    
    </tr>
  </table>
<table cellpadding="2" cellspacing="1"  width="95%" class="toptable grid" border="1" align=center>
  <form action="cwlb_add2.asp?id=<%=ID%>&action=add" method=post id=form1 name=form1 onSubmit="return chk(this)">

    <tr bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">标志：</div></td>
      <td width="7%" class="category"> 
      	<select name="SCORE" id="SCORE">
          <%
   Response.write "<option value='借'>借</option>"
   Response.write "<option value='贷'>贷</option>"

%>
        </select>
      	
      	</td>
      <td width="8%" class="category"><div align="right">科目代码：</div></td>
      <td width="7%" class="category"> <input name="XJCODE" type="text" id="XJCODE"  size="10" maxlength="10" onKeyPress="javascript:CheckNum();"></td>
      <td width="10%" class="category"><div align="right">科目名称：</div></td>
      <td colspan="2" class="category"><input name="XJNAME" type="text" id="XJNAME"  size="12" maxlength="12"> 

    </tr>
    <tr  bgcolor=ffffff> 

      <td width="8%" class="category"><div align="right">银行名称：</div></td>
      <td width="7%" class="category"> <input name="XJYH" type="text" id="XJYH"  size="10" maxlength="10"></td>

      <td width="8%" class="category"><div align="right">银行账号：</div></td>
      <td width="10%" class="category"> <input name="YHZH" type="text" id="YHZH"  size="20" maxlength="150"></td>

      <td width="8%" class="category"><div align="right">联系人：</div></td>
      <td width="7%" class="category"> <input name="YHLXR" type="text" id="YHLXR"  size="10" maxlength="10"></td>


    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">联系电话：</div></td>
      <td width="7%" class="category"> <input name="XGDH" type="text" id="XGDH"  size="10" maxlength="10"></td>
      <td width="8%" class="category"><div align="right">类别：</div></td>
      <td width="7%" class="category"> <input name="LB" type="text" id="LB"  size="10" maxlength="10"></td>

      <td  width="8%" class="category"><div align="right">备注：</div></td>
      <td width="17%" class="category"> <input name="remark" type="text" id="remark" " size="20" maxlength="50"></td>



    </tr>
    <tr  bgcolor=ffffff> 
      <td width="8%" class="category"><div align="right">期初金额：</div></td>
      <td width="7%" class="category"> <input name="qc" type="text" id="qc"  size="10" maxlength="10" onKeyPress="javascript:CheckNum();"></td>
      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>

      <td width="8%" class="category"><div align="right"></div></td>
      <td width="7%" class="category"> </td>
    <tr  bgcolor=ffffff> 
      <td colspan="7" align="center" class="category"><input type="submit" name="Submit" value="添 加"></td>
    </tr>
  </form>
</table>
<%


End if%>
</html>