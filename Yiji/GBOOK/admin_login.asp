<!--#include file="conn.asp"-->
<!--#include file="pageset.asp"-->
<!--#include file="lock.asp"-->
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<%
if request("Submit")<>"" then
   textfield=request("textfield")
   textfield2=request("textfield2")
   if instr(textfield2,"'")<>0 or instr(textfield,"'")<>0 then call errmsg("�����к��зǷ��ַ�")
   sql="select username,password from admin where username='"&textfield&"' and password='"&textfield2&"'"
   rs.open sql,conn,1,3
   if rs.eof and rs.bof then
   call errmsg("�û��������벻��ȷ")
   else
   session("admin")=rs("username")
   session("pwd")=rs("password")
   response.redirect"jlsk.asp"
   end if
   rs.close
   call connclose
response.end
end if
%>
<body  topmargin="0">
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td> </td>
  </tr>
  <tr> 
    <form name="form1" method="post" action="">
      <td height="356"> 
        <table width="32%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="26"><div align="center"><b>����Ա��¼</b></div></td>
          </tr>
          <tr> 
            <td height="22">�û����� 
              <input type="text" name="textfield"></td>
          </tr>
          <tr> 
            <td height="26">��&nbsp;&nbsp;�룺 
              <input type="password" name="textfield2"></td>
          </tr>
          <tr> 
            <td height="40"><div align="center"> 
                <input type="submit" name="Submit" value="��¼">
                <input type="reset" name="Submit2" value="����">
              </div></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center" height="40">
  <tr> 
  </tr>
</table>
</body> 
</html>

