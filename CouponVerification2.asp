<!--#include file="include/SQLConn.inc" -->
<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
'response.write UserIPAddress

Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station , LoginUser , Outdoor From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

StationID = Rs1("Station")

CarPlate  = Replace(Request("CarPlate")," ","")

ProductType = Request("ProductType")

If ProductType = "" then

   ProductType = "53"

End If

Message = ""


'Response.write "Station ID:  " & StationID & "<br/>"
'Response.write "User ID:  " & Rs1("LoginUser")   & "<br/>"
%>

<HTML>
<HEAD>
<title>
 - 禮券驗證系統  -
</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link href="include/css/bootstrap.min.css" rel="stylesheet" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
<SCRIPT language=JavaScript>
<!--
function check()
{
document.barCodeForm.Barcode.focus();
}

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
    padding: 30px 80px;
    background-color: #9fdf9f;
    text-align: center;
    text-shadow: 0 1px 1px rgba(255,255,255,0.75);
    vertical-align: middle;
    cursor: pointer;
    border: 1px solid #ccc;
   
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

</style>
</HEAD>
<body class="homepage" onload="document.barCodeForm.Barcode.focus();">

<!--#include file="include/header.inc" -->

<form name="barCodeForm" method="post" action="" onSubmit="javascript:return mod10CheckDigit();">

<h1 class="Title">禮券驗證系統</h1> 

<div align="right"><a href="DailyReport.asp">驗證紀錄</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="logout.asp">登出</a></div>


<div align="center">

<table width="80%" border="0"  cellpadding="15" cellspacing="4" class="Ver">


     <tr>

	    <td width="25%">產品類型

        </td>

		<td width="25%">

         <input type="radio" id="radio1" name="ProductType" value="53" <% If ProductType = "53" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio1" >無鉛</label>
        </td> 

	    <td width="25%">

         <input type="radio" id="radio2" name="ProductType" value="52" <% If ProductType = "52" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio2">油渣</label>
        </td> 

	    <td width="25%">

         <input type="radio" id="radio3" name="ProductType" value="54" <% If ProductType = "54" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio3">V 能量</label>
     </tr>


     <tr>

	    <td width="25%"></td>

	   

	    <td width="25%">

         <input type="radio" id="radio5" name="ProductType" value="LB" <% If ProductType = "LB" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio5">機油</label>
        </td> 

	    <td width="25%">

         <input type="radio" id="radio6" name="ProductType" value="CW" <% If ProductType = "CW" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio6">洗車</label>

        </td> 

        <td width="20%">

         <input type="radio" id="radio4" name="ProductType" value="CR" <% If ProductType = "CR" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">

          <label for="radio4">便利店</label>
        </td>

        </tr>


    <tr>

	    <td>禮券號碼 

        </td> 

		<td >

					
<input name="Barcode" type="text" autocomplete="off" value="" size="20" maxlength="15" id="bc" ">

         </td>

         <td colspan="2">

<input name="StationID" type="hidden" value="<%=StationID%>">

	<input type="Submit" Name="Button" value="確定">


        </td>

     </tr>


</table>

</div>

</form>

<div class="Message">

 <%

        Message = Request("Message")
       
      
        If Message <> "" then


           Response.write Message


        End If

        Message = ""
       

  %>



</div>

<div align="right">

<table width="100%" border="0"  cellpadding="15" cellspacing="4" class="Login">


     <tr>

	    <td align="right">

           <font size="-1">050072137000120     31 Nov 2015 12:20 </font><br>
           <font size="-1">050072137000121     31 Nov 2015 12:18 </font><br>
<font size="-1">050072137000114     31 Nov 2015 12:16 </font><br>
<font size="-1">050072137000113     31 Nov 2015 11:20 </font><br>
<font size="-1">050072137000112     31 Nov 2015 10:25 </font><br>


        </td>

     </tr>

</table>



</div>

</body>
</HTML>
