
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
<title><%=dianming%> - ��ӡ������ͬ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style2.css" rel="stylesheet" type="text/css">
<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
window.document.body.innerHTML=bdhtml; 
         }
</script>
</HEAD>

<BODY>
<%
  set rs_contract=conn.execute("select * from salescontract where contractnum="&request("contractnum"))
  sql="select * from owncompany where company="&rs_contract("owncompany")
  set rs_owncompany=conn.execute(sql)  
%>


<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td height="21" align="center"><img src="../images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();window.close()"></td>
  </tr>
</table>
<!--startprint-->

<!--logo-->
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td width="50%" height="100" align="left"><img src="../images/qz.jpg" align="absmiddle"></td>
    <td width="50%">
      <table border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
        <tr>
          <td width="10%" align="left">��ַ��</td>
          <td width="90%" align="left"><%=rs_owncompany("fullname")%></td>
        </tr>
        <tr>
          <td width="10%" align="left">�绰��</td>
          <td width="90%" align="left"><%=rs_owncompany("tel")%></td>
        </tr>    
        <tr>
          <td width="10%" align="left">���棺</td>
          <td width="90%" align="left"><%=rs_owncompany("fax")%></td>
        </tr>
        <tr>
          <td width="10%" align="left">���䣺</td>
          <td width="90%" align="left"><%=rs_owncompany("email")%></td>
        </tr>        
      </table>
    </td>
  </tr>
</table>

<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <td width="100%" align="center">
      <h2>
        <%if rs_contract("category")="A" then %>
          ������ͬ
        <%else%>
          �ڻ�������ͬ
        <%end if%>
      </h2>
    </td>
  </tr>
</table>
<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <th width="15%" height="30" align="left">�׷�����������</th>
	  <td width="35%" height="30" align="left">
          <%
          sql="select * from owncompany where company="&rs_contract("owncompany")
          set rs_owncompany=conn.execute(sql)
          %>
          <%=rs_owncompany("fullname")%>
    </td>	  
	  <th width="35%" height="30" align="right">��ͬ��ţ�</th>	
	  <td width="15%" height="30" align="left"><%=rs_contract("contractnum")%></td>
	</tr>		
  <tr> 
    <th width="15%" height="30" align="left">�ҷ����򷽣���</th>
    <td width="35%" height="30" align="left">
          <%
          sql="select * from customer where customername="&rs_contract("customer")
          set rs_customer=conn.execute(sql)
          %>
          <%=rs_customer("fullname")%>
    </td>   
    <th width="35%" height="30" align="right">��˾���棺</th>  
    <td width="15%" height="30" align="left"><%=rs_owncompany("fax")%></td>
  </tr> 
</table>

<p style="margin-left:20pt;">
  �ּ���˫�������Ѻ�Э�̡����ݻ�����ԭ��ͬ�ⰴ����������ǩ������ͬ��
</p>

<table width="94%" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000" style="margin-top:18pt;margin-left:20pt;">
  <tr>
  <th>��˾���</th>
  <th>����</th>
	<th>����</th>
	<th>Ʒ��</th>
  <th>���</th>  
  <th>��װ</th>  
  <%if rs_contract("category")="A" then %>
    <th>����</th>
  <%end if%>  
	<th>�������֣�</th>
	<th>���ۣ�Ԫ/�֣�</th>
  <%if rs_contract("category")="A" then %>
    <th>������</th>
    <th>������</th>
  <%else%>
    <th>Ԥ�Ƶ�����</th>
    <th>������</th>      
  <%end if%>  
  </th>	
  </tr>

  <tr>
  	<td align="center"><%=rs_contract("refshipment")%></td>
  	<td align="center"><%=rs_contract("country")%></td>
  	<td align="center"><%=rs_contract("plant")%></td>
  	<td align="center"><%=rs_contract("material")%></td>
  	<td align="center"><%=rs_contract("spec")%></td>
  	<td align="center"><%=rs_contract("package")%></td>
    <%if rs_contract("category")="A" then %>
      <td align="center"><%=rs_contract("quantity")%></td>
    <%end if%>  
  	<td align="center"><%=rs_contract("weight")%></td>
  	<td align="center"><%=rs_contract("price")%></td>
    <%if rs_contract("category")="A" then %>
      <td align="center"><%=rs_contract("coldstorage")%></td>
      <td align="center"><%=rs_contract("deliveryloc")%></td>
    <%else%>
      <td align="center"><%=rs_contract("boarddate")%></td>
      <td align="center"><%=rs_contract("deliveryport")%></td>     
    <%end if%>  
  </tr>				
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>   
    <td align="center">�ܼ�(Ԫ):</td>
    <td align="center"><%=rs_contract("totalamount")%></td>
    <td></td>  
    <td></td>  
  </tr>     
</table>

  <br/>
  <p>
    ��ע���˹�Ĺ��Ϊ�� <%=rs_contract("case")%>   ���ɸ��ݹ�����в�ѯ���ڡ�
  </p>
  <br/>
  <p>
    �����˺ţ�
    <br/>
      �˻����ƣ� <%=rs_owncompany("bankaccountname")%>
      �˺ţ� <%=rs_owncompany("bankaccount")%>
      �����У� <%=rs_owncompany("bankname")%>
    <% if rs_owncompany("bankaccountname2")<>"" then%>
      <br/>
      �˻����ƣ� <%=rs_owncompany("bankaccountname2")%>
      �˺ţ� <%=rs_owncompany("bankaccount2")%>
      �����У� <%=rs_owncompany("bankname2")%>
    <%end if%>
  </p>  
  <br/>	
  <p>
    ���ʽ������֧����������48Сʱ��ǩ��˺�ͬ���ش��׷�������2����������֧������ RMB 30000Ԫ��ʣ����� RMB <% remain=rs_contract("totalamount")-30000%><%=remain%>Ԫ�ڻ��ﵽ��ǰ<%if rs_contract("category")="A" then %>һ<%else%>��<%end if%>�츶�塣��ʵ�ʹ�Ӧ��װ��ʱ�䡢װ������Ϊ׼��������ʵ�ʵ�����������Ӧ���������򷽲������ڻش�����ͬ��֧������������Ȩ�ӵ�ʱ�г�״̬������ȡ���˺�ͬ��
  </p>
  <br/>
  <p>
    �������ڣ������̼캣�������������̣����ۺ�������������ҿ��������������������ǩ���ġ�����֤��Ϣ��������⣬��ط��ɼ׷�֧������������˳�ӵ������̼�����ա�
  </p>
  <br/>
  <p>
    ������ʽ���ҷ����ᡣ<%if rs_contract("category")="B" then %>����ԭ�񵹳������������ҷ��е���<%end if%>
  </p>
  <br/>
  <p>
    ����˵�����������Ϊҵ��ȷ�ϵ����ݣ������ﵽ��ʱ���ͻ��޷��굥����ͻ����ге��г����֮��ʧ��
  </p>
  <br/>
  <p>
    ����������յ���Ʒ���緢����������Ʒ������״�����⣬���ڵ���48Сʱ����������������������⣬���ڵ���7����������������δ��˫��ͬ�⣬���ôӻ��������п۳����ҷ�������ṩ�꾡��������������Ƭ����������˵���ȵȣ����б�Ҫ˫�������������֤���ؽ��й�֤����֤�������η��е�����˾������δ���ġ��𻵰�װ�Ĳ�Ʒ������������������˫��Э�̺���Ҫ�����˻������ˣ���˾ֻ����ԭ��װ�˻�����
  </p>
  <br/>
  <p>
    ������������˫���������г�������䶯��ΥԼ��������һ���������й�׷��Ȩ��������ս�������𣬰չ��Ȳ��ɿ���������ɵ��޷���Լ���ú�ͬ�Զ�ʧЧ���׷��˻���ͬ��������δ�����ˣ�˫��Э�̽����
    ���򱾺�ͬ����Ļ��뱾��ͬ�йص�һ�����飬����˫����Ӧ����ͬ���涨Э�̽����Э�̲��������Ϻ����ٲ�ίԱ�������ٲã���ط����ɰ��߷��е���
  </p>
  <br/>
  <p>
    ����ͬһʽ���ݣ�����˫����ִһ�ݡ���˫����Ȩ����ǩ�ֻ���²��յ�����֮������Ч����ӡ�������������ͬ������Ч����
  </p>
  <br/>
  <p bold>
    �������ѣ��ڻ����ڻ��в��������ܻ�˳��1-2�ܣ���˾���չ涨��ÿ�ܻ���ͻ��ṩ���ڱ���δ�յ��뼰ʱ��ϵ�Ա��������·��ͣ�����˾���͸��Ŵ��ڱ�������������ڣ��ͻ�δ�������ģ���ͬ�ͻ��Ͽɴ��ڱ����ݡ�
  </p>

<table width="94%" border="0" cellpadding="0" cellspacing="2" align="center" style="margin-top:18pt;margin-left:20pt;">
  <tr> 
    <th width="15%" height="30" align="left">�׷���</th>
    <td width="35%" height="30" align="left"><%=rs_owncompany("fullname")%></td>   
    <th width="35%" height="30" align="left">�ҷ���</th>  
    <td width="15%" height="30" align="left"><%=rs_customer("fullname")%></td>
  </tr>   
  <tr> 
    <th width="15%" height="30" align="left">�׷���Ȩ��</th>
    <td width="35%"></td>   
    <th width="35%" height="30" align="left">�ҷ���Ȩ��</th>  
    <td width="15%"></td>
  </tr> 
  <tr> 
    <th width="15%" height="30" align="left">���ڣ�</th>
    <td width="35%" height="30" align="left">date()</td>   
    <th width="35%" height="30" align="left">���ڣ�</th>  
    <td width="15%"></td>
  </tr>   
</table>

<!--endprint-->
</body>
</html>