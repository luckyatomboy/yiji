

  //窗口表格增加一行
  function addNewRow(){
   var tabObj=document.getElementById("item");//获取添加数据的表格
   var rowsNum = tabObj.rows.length;  //获取当前行数
   var colsNum=tabObj.rows[0].cells.length;//获取行的列数
   var myNewRow = tabObj.insertRow(rowsNum);//插入新行.
      
   var itemObj=document.getElementsByName("itemno");//取得所有行的itemno
   var itemNo=1;
   if (itemObj.length==0) {
  	 itemNo=1;//如果没有item,给1
  	 
  }else{
  	 itemNo=parseInt(itemObj[itemObj.length-1].value) + 1; //取最大行itemno加1
  } 
   var newTdObj1=myNewRow.insertCell(0);
   newTdObj1.innerHTML="<input type='checkbox' name='chkArr'  id='chkArr' align='middle'/>";
   var newTdObj2=myNewRow.insertCell(1);
   newTdObj2.innerHTML="<input type='text' name='itemno' id='itemno' style='width:30px' value='"+itemNo+"' readonly='true'/>"; 
   var newTdObj3=myNewRow.insertCell(2);
	 newTdObj3.innerHTML="<select name='material"+itemNo+"' id='material"+itemNo+"' style='width:100px'>"
	    +" <% for i = 0 to materialCount-1 %>"
	 		+" <option value='<%=material(i)%>'><%=material(i)%></option>"
	 		+" <% next %> </select>";
   var newTdObj4=myNewRow.insertCell(3);
//   newTdObj4.innerHTML="<input type='text' name='customer"+itemNo+"' id='customer"+itemNo+"' />";
	 newTdObj4.innerHTML="<select name='customer"+itemNo+"' id='customer"+itemNo+"' style='width:100px'>"
	    +" <% for i = 0 to customerCount-1 %>"
	 		+" <option value='<%=customer(i)%>'><%=customer(i)%></option>"
	 		+" <% next %> </select>";   
   var newTdObj5=myNewRow.insertCell(4);
   newTdObj5.innerHTML="<input type='text' name='spec"+itemNo+"' id='spec"+itemNo+"' style='width:100px'/>";   
   var newTdObj6=myNewRow.insertCell(5);
   newTdObj6.innerHTML="<input type='text' name='package"+itemNo+"' id='package"+itemNo+"' style='width:100px'/>";      
   var newTdObj7=myNewRow.insertCell(6);
   newTdObj7.innerHTML="<input type='text' name='contractWeight"+itemNo+"' id='contractWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj8=myNewRow.insertCell(7);
   newTdObj8.innerHTML="<select name='contractWeightUom"+itemNo+"' id='contractWeightUom"+itemNo+"' style='width:50px'>"
      +" <option value='公斤'>公斤</option>"
      +" <option value='磅'>磅</option>"
      +" </select>";       
   var newTdObj9=myNewRow.insertCell(8);
   newTdObj9.innerHTML="<input type='text' name='actualWeight"+itemNo+"' id='actualWeight"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj10=myNewRow.insertCell(9);
   newTdObj10.innerHTML="<input type='text' name='purchasePrice"+itemNo+"' id='purchasePrice"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj11=myNewRow.insertCell(10);
   newTdObj11.innerHTML="<select name='priceunit"+itemNo+"' id='priceunit"+itemNo+"' style='width:50px'>"
      +" <option value='公斤'>公斤</option>"
      +" <option value='磅'>磅</option>"
      +" </select>";      
   var newTdObj12=myNewRow.insertCell(11);
   newTdObj12.innerHTML="<input name='produceDate"+itemNo+"' id='produceDate"+itemNo+"' style='width:80px' class='datepicker'>";
//   	+ " <img src='../images/date.gif' align='absmiddle' style='cursor:pointer;'" 
//   	+ " onClick=\""+"JavaScript:window.open('../day.asp?form=formitem&field=produceDate"+itemNo+"&oldDate=produceDate"+itemNo+".value','','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=250,height=170,top=150,left=590');\"/>";     
   var newTdObj13=myNewRow.insertCell(12);   	//箱数
	 newTdObj13.innerHTML="<input name='casenum"+itemNo+"' id='casenum"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   	
   var newTdObj14=myNewRow.insertCell(13);   	//发票总金额
	 newTdObj14.innerHTML="<input name='invAmount"+itemNo+"' id='invAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 
   var newTdObj15=myNewRow.insertCell(14);
   newTdObj15.innerHTML="<select name='invCurr"+itemNo+"' id='invCurr"+itemNo+"' style='width:50px'>"
	 		+" <option value='美元'>美元</option>"
	 		+" <option value='澳币'>澳币</option>"
	 		+" </select>";   
   var newTdObj16=myNewRow.insertCell(15);   	//尾款金额
	 newTdObj16.innerHTML="<input name='finalAmount"+itemNo+"' id='finalAmt"+itemNo+"' style='width:80px' value=0 onKeyPress=\""+"javascript:CheckNum();\""+" onKeyUp=\""+"this.value=this.value.replace(/[^\\d.]/g,'')\""+">";   		 

  }

//窗口表格删除一行
  function removeRow(){
   var chkObj=document.getElementsByName("chkArr");
   var tabObj=document.getElementById("item");
   for(var k=0;k<chkObj.length;k++){
    if(chkObj[k].checked){
     tabObj.deleteRow(k+1);
     k=-1;
    }
   }
  }

function check1(){
<!-- 检查合同号 -->
if (document.form1.contract.value=="")
{
alert("请输入合同号！");
return false;
}
<!-- 检查自动证 -->
<!-- 检查自动证 -->
}

function chsel(vendor){
//设置国家	
    document.form1.country.length = 0; 
    for (i=0; i<country.length; i++) 
    { 
        if (country[i][0]==vendor) 
        {document.form1.country.options[0] = new Option(country[i][1]);} 
    } 
//设置工厂    
    document.form1.plant.length = 0; 
    for (m=0; m<plant.length; m++) 
    {    	
      if (plant[m][0] == vendor) 
      {
      	for (n=1;n<5;n++)
      	{
      		if (plant[m][n]!='') {
      			document.form1.plant.options[document.form1.plant.length] = new Option(plant[m][n]);
      		}
      	} 
    	}  
    }       
//设置付款条件   
    document.form1.incoterm.length = 0; 
    for (m=0; m<incoterm.length; m++) 
    {    	
      if (incoterm[m][0] == vendor) 
      {
      	for (n=1;n<3;n++)
      	{
      		if (incoterm[m][n]!='') {
      			document.form1.incoterm.options[document.form1.incoterm.length] = new Option(incoterm[m][n]);
      		}
      	} 
    	}  
    }       
}
   