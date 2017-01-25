<!--#include file="include/SQLConn.inc" -->
<%

Message = Request("Message")

Dim UserIPAddress
UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
UserIPAddress = "192.168.1.1"
'response.write UserIPAddress



Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

If Not Rs1.EoF Then

StationID = Rs1("Station")

Else 

Alert = "油站編號未設置, 請提供此電腦的 IP 地址 [ " & UserIPAddress & " ] 給系統管理員"


End If

%>

<HTML>
<HEAD>
<title>
    --  禮券驗證系統  -- 
</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<meta name="viewport" content="user-scalable=no" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
<style type="text/css">
.btns
{  
width: 100px;
height:60px; 
}
</style>

<script type="text/javascript">

onload=function(){ attachHandlers(); }

function attachHandlers(){
  var the_nums = document.getElementsByName("number");
  var the_nums2 = document.getElementsByName("number2");
  for (var i=0; i < the_nums.length; i++) { the_nums[i].onclick=inputNumbers; }

  for (var i=0; i < the_nums2.length; i++) { the_nums2[i].onclick=inputNumbers2; }
}

function inputNumbers() {
  var the_field = document.getElementById('UserID');
  var the_value = this.value;
  switch (the_value) {
    case '<' :
      the_field.value = the_field.value.substr(0, the_field.value.length - 1);
      break;
    default : document.getElementById("UserID").value += the_value;
      break;
  }
  document.getElementById('UserID').focus();
  return true;
}

function inputNumbers2() {
  var the_field2 = document.getElementById('Password');
  var the_value2 = this.value;
  switch (the_value2) {
    case '<' :
      the_field2.value = the_field2.value.substr(0, the_field2.value.length - 1);
      break;
    default : document.getElementById("Password").value += the_value2;
      break;
  }
  document.getElementById('Password').focus();
  return true;
}

function ShowKey1(){   
   var the_keypad1 = document.getElementById('keypad1');
   var the_keypad2 = document.getElementById('keypad2');
   the_keypad1.style.display ="block";
   the_keypad2.style.display ="none";

}

function ShowKey2(){   
   var the_keypad1 = document.getElementById('keypad1');
   var the_keypad2 = document.getElementById('keypad2');
   the_keypad2.style.display ="block";
   the_keypad1.style.display ="none";

}

</script>
</HEAD>
<body class="homepage">
<!--#include file="include/header.inc" -->

<form name="fm1" method="post" action="">

<div align="center">

<h1 class="Title">禮券驗證系統</h1>

<table width="60%" border="0" cellpadding="10" cellspacing="0" class="Login">

    <tr>

	    <td width="40%">油站編號 

        </td> 

		<td colspan="2">
					<% = StationID %>
        
        </td>

     </tr>

     <tr>

	    <td width="40%">用戶編號

        </td> 

		<td >
<%If Alert = "" Then%>
			<input name="UserID" ID="UserID" type=text value="" class="ver" size="10"  autocomplete="off" onfocus="ShowKey1()">
<%End If%>
        </td>

        <td>

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
     <input type="button" name="number" value="." id="_dot" class="btns">
     <input type="button" name="number" value="0" id="_0" class="btns">
     <input type="button" name="number" value="<" id="chgsign" class="btns">
    </div>
   </div>
        </td>

     </tr>

     <tr>

	    <td width="40%">用戶密碼

        </td> 

		<td >

			<input name="Password" ID="Password" class="ver" type="password" value="" size="10" autocomplete="off" <%If Alert<>"" Then%>Readonly<%End If%> onfocus="ShowKey2()">

        </td>

      <td>

    <div id="keypad2" style="width:135px;float:left;display:none;">
 
    <div id="row1">
     <input type="button" name="number2" value="1" id="_1" class="btns">
     <input type="button" name="number2" value="2" id="_2" class="btns">
     <input type="button" name="number2" value="3" id="_3" class="btns">
    </div>
    <div id="row2">
     <input type="button" name="number2" value="4" id="_4" class="btns">
     <input type="button" name="number2" value="5" id="_5" class="btns">
     <input type="button" name="number2" value="6" id="_6" class="btns">
    </div>
    <div id="row3">
     <input type="button" name="number2" value="7" id="_7" class="btns">
     <input type="button" name="number2" value="8" id="_8" class="btns">
     <input type="button" name="number2" value="9" id="_9" class="btns">
    </div>
    <div id="row4">
     <input type="button" name="number2" value="." id="_dot" class="btns">
     <input type="button" name="number2" value="0" id="_0" class="btns">
     <input type="button" name="number2" value="<" id="chgsign" class="btns">
    </div>
   </div>
        </td>
     </tr>

     <tr>

	    <td width="40%">

        </td> 

		<td colspan="2"><input Type="Button" Name=" 確定 " value="   確定   " onClick="doLogin();" style="background: #9fdf9f;border:1px solid #00ff00;font-size:40px">
             <input name="StationID" type="hidden" value="<%=StationID%>">


        </td>

     </tr>


    


</table>

</div>

</form>

<br/><br/><br/>

<div class="Message">

 <%

        If Message = "Fail" then

           Response.write "用戶編號或密碼不正確"

        End If

        response.write  Alert

  %>

</div>

</body>
</HTML>
