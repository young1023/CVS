<!--#include file="include/SQLConn.inc" -->
<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
'response.write UserIPAddress

Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station , LoginUser From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)



ProductType = Request("ProductType")

If ProductType = "" then

   ProductType = "54"

End If



%>

<HTML>
<HEAD>
<title>
 - 禮券驗證系統  -
</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<meta name="viewport" content="user-scalable=no" />
<link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
<SCRIPT language=JavaScript>
<!--
onload=function(){ attachHandlers(); }

function attachHandlers(){
  var the_nums = document.getElementsByName("number");
  for (var i=0; i < the_nums.length; i++) { the_nums[i].onclick=inputNumbers; }
}

function inputNumbers() {
  var the_field = document.getElementById('bc');
  var the_value = this.value;
  switch (the_value) {
    case '<' :
      the_field.value = the_field.value.substr(0, the_field.value.length - 1);
      break;
    default : document.getElementById("bc").value += the_value;
      break;
  }
  document.getElementById('bc').focus();
  return true;
}


function ShowKey1(){   
   var the_keypad1 = document.getElementById('keypad1');
   the_keypad1.style.display ="block";

}



function stopEnterKey(evt) {
        var evt = (evt) ? evt : ((event) ? event : null);
        var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
        if ((evt.keyCode == 13) && (node.type == "text")) { return false; }
    }
    document.onkeypress = stopEnterKey;


//-->
</SCRIPT>
<style>
/*
  Hide radio button (the round disc)
  we will use just the label to create pushbutton effect
*/
input[type=radio] {
    display:none; 
    margin:10px;
}
 
/*
  Change the look'n'feel of labels (which are adjacent to radiobuttons).
  Add some margin, padding to label
*/
input[type=radio] + label {
    display:inline-block;
    margin: -2px;
    padding: 10px 80px;
    background-color: #9fdf9f;
    text-align: center;
    text-shadow: 0 1px 1px rgba(255,255,255,0.75);
    vertical-align: middle;
    cursor: pointer;
    border: 2px solid #ccc;   
    border-color: #e6e6e6 #e6e6e6 #bfbfbf;
    border-color: rgba(0,0,0,0.1) rgba(0,0,0,0.1) rgba(0,0,0,0.25);
    border-bottom-color: #b3b3b3;
   
}
/*
 Change background color for label next to checked radio button
 to make it look like highlighted button
*/
input[type=radio]:checked + label { 
    background-image: none;
    background-color:#ffa31a;
    border-color: #cc7a00;
}

.btns
{  
width: 30px;
height:10px; 
}

</style>
</HEAD>
<body class="homepage" onload="document.barCodeForm.Barcode.focus();">

<!--#include file="include/header.inc" -->




<div align="center">
<a href="DailyReport.asp">驗證紀錄</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="logout.asp">登出</a></div>


<div align="center">

<form name="barCodeForm" method="post" action="" onSubmit="javascript:return mod10CheckDigit();">

<table width="80%" border="0"  cellpadding="10" cellspacing="0" class="Ver">


     <tr>

       <td width="30%" >

         <input type="radio" id="radio1" name="ProductType" value="54" <% If ProductType = "54" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio1">V  能量</label>
        </td>

		<td width="30%" >

         <input type="radio" id="radio2" name="ProductType" value="53" <% If ProductType = "53" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio2" >無鉛</label>
        </td> 

	    <td width="30%" >

         <input type="radio" id="radio3" name="ProductType" value="52" <% If ProductType = "52" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio3">柴油</label>


        </td> 

  
	 
     </tr>



     <tr >

        <td >

      
         <input type="radio" id="radio4" name="ProductType" value="CR" <% If ProductType = "CR" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">

          <label for="radio4">便利店</label>

        

        </td>
    	   

	    <td >

         <input type="radio" id="radio5" name="ProductType" value="LB" <% If ProductType = "LB" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio5">機油</label>
        </td> 

	    <td >

         <input type="radio" id="radio6" name="ProductType" value="CW" <% If ProductType = "CW" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio6">洗車</label>

        </td> 

        </tr>


    <tr>

	    <td>禮券號碼 

        </td> 

		<td >

					
<input name="Barcode" class="ver" type="text" autocomplete="off" value="" size="18" maxlength="15" id="bc" onkeypress="ShowKey1()">

         </td>

         <td >

 
	<input type="Submit" Name="Button" value="   確定   " style="background: #9fdf9f;border:1px solid #00ff00;font-size:46px">



    
<div id="keypad1" style="width:135px;float:left;display:none;">
 
    <div id="row1">
     <input type="button" name="number" value="1" id="_1" class="btns">
     <input type="button" name="number" value="2" id="_2" class="btns">
     <input type="button" name="number" value="3" id="_3" class="btns">
    </div>
    <div id="row2">
     <input type="button" name="number" value="4" id="_4" class="btns">
     <input type="button" name="number" value="5" id="_5" class="btns">
     <input type="button" name="number" value="6" id="_6" class="btns">
    </div>
    <div id="row3">
     <input type="button" name="number" value="7" id="_7" class="btns">
     <input type="button" name="number" value="8" id="_8" class="btns">
     <input type="button" name="number" value="9" id="_9" class="btns">
    </div>
    <div id="row4">
     <input type="button" name="number" value="0" id="_0" class="btns">
     <input type="button" name="number" value="<" id="chgsign" class="btns">
    </div>
   </div>

        </td>

     </tr>


</table>



</form>


<% 

   Color = request("Color")

   If Color = 1 then

       bgcolor = "#9fdf9f"

   Elseif Color = 2 then

       bgcolor = "#ffa31a"

   Else

       bgcolor = "#ffffff"

   End if

%>

<table width="80%" border="0"  cellpadding="30" cellspacing="2" bgcolor="<%=bgcolor%>" class="Message">


     <tr>

	    <td align="center">






 <%

        Message = Request("Message")
       
      
        If Message <> "" then

%>

 



<%


        Response.write Message


        End If

        Message = " "
       

  %>

        </td>

	
     </tr>

</table>





</div>


</body>
</HTML>
