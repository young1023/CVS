<!--#include file="include/SQLConn.inc" -->
<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
'response.write UserIPAddress

Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station , LoginUser From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

StationID = Rs1("Station")

ProductType = Request("ProductType")

If ProductType = "" then

   ProductType = "54"

End If



'Response.write "Station ID:  " & StationID & "<br/>"
'Response.write "User ID:  " & Rs1("LoginUser")   & "<br/>"
%>

<HTML>
<HEAD>
<title>
 - §�����Ҩt��  -
</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
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
    padding: 13px 80px;
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


</style>
</HEAD>
<body class="homepage" onload="document.barCodeForm.Barcode.focus();">

<!--#include file="include/header.inc" -->

<form name="barCodeForm" method="post" action="" onSubmit="javascript:return mod10CheckDigit();">

<h2 class="Title">§�����Ҩt��</h2> 

<div align="right"><a href="DailyReport.asp">���Ҭ���</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="logout.asp">�n�X</a></div>


<div align="center">

<table width="90%" border="0"  cellpadding="6" cellspacing="2" class="Ver">


     <tr>

       <td width="25%">

         <input type="radio" id="radio1" name="ProductType" value="54" <% If ProductType = "54" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio1">V  ��q</label>
        </td>

		<td width="25%">

         <input type="radio" id="radio2" name="ProductType" value="53" <% If ProductType = "53" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio2" >�L�]</label>
        </td> 

	    <td width="25%">

         <input type="radio" id="radio3" name="ProductType" value="52" <% If ProductType = "52" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio3">��o</label>
        </td> 

        <td width="25%">

    	<input type="Submit" Name="Button" value=" ��L�ﶵ  " style="background: #9fdf9f;border:1px solid #00ff00;font-size:36px" onclick="NextPage1()">

      </td> 
	 
     </tr>


     <tr>

        <td >

         <input type="radio" id="radio4" name="ProductType" value="CR" <% If ProductType = "CR" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">

          <label for="radio4">�K�Q��</label>
        </td>
    	   

	    <td >

         <input type="radio" id="radio5" name="ProductType" value="LB" <% If ProductType = "LB" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio5">���o</label>
        </td> 

	    <td >

         <input type="radio" id="radio6" name="ProductType" value="CW" <% If ProductType = "CW" Then%>Checked<%End If%> OnClick="document.barCodeForm.Barcode.focus();">
         <label for="radio6">�~��</label>

        </td> 

   

        

        </tr>


    <tr>

	    <td>§�鸹�X 

        </td> 

		<td >

					
<input name="Barcode" class="ver" type="text" autocomplete="off" value="" size="18" maxlength="15" id="bc" ">

         </td>

         <td colspan="2" align="left">

    <input name="StationID" type="hidden" value="<%=StationID%>">

	<input type="Submit" Name="Button" value="  �T�w  " style="background: #9fdf9f;border:1px solid #00ff00;font-size:46px">


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
