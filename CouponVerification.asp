<!--#include file="include/SQLConn.inc" -->
<%



UserIPAddress = Request.ServerVariables("REMOTE_ADDR")

Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select Station , LoginUser From Station Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

StationID = Rs1("Station")

response.write stationid

CarPlate  = Replace(Request("CarPlate")," ","")

ProductType = Request("ProductType")

Message = ""


'Response.write "Station ID:  " & StationID & "<br/>"
'Response.write "User ID:  " & Rs1("LoginUser")   & "<br/>"
%>

<HTML>
<HEAD>
<title>
 - §�����Ҩt�� (�D��) -
</title>
<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
</HEAD>
<% If ProductType = "" Then %>
<body class="homepage" onload="document.barCodeForm.CarPlate.focus();">
<% Else %>
<body class="homepage" onload="document.barCodeForm.Barcode.focus();">
<% End If %>
<!--#include file="include/header.inc" -->

<form name="barCodeForm" method="post" action="" onSubmit="javascript:return mod10CheckDigit();">

<div align="center">

<h1 class="Title">§�����Ҩt�� (�D��)</h1>


<table width="60%" border="0" cellpadding="04" cellspacing="0" class="Report">

      <tr>

              <td>

                      <a href="DailyReport.asp">���Ҭ���</a>                 

              </td>

               <td>

                      <a href="logout.asp">�n�X</a>                 

              </td>

      </tr>

</table>

<br/>



<table width="60%" border="0"  cellpadding="04" cellspacing="2" class="Login">

      <tr>

        <td width="23%">���P

        </td> 

		<td width="37%">

					<input name="CarPlate" type=text value="<%=CarPlate%>" size="20" Maxlength = "20" autocomplete="off">



        </td>


     </tr>
    

     <tr>

	    <td width="23%">���~����

        </td> 

		<td width="37%">

                    <select name="ProductType" id="ProductType">

                    <option value="53" <% If ProductType = "53" Then%>Selected<%End If%>>�L�]</option>

                    <option value="52" <% If ProductType = "52" Then%>Selected<%End If%>>�o��</option>

                    <option value="54" <% If ProductType = "54" Then%>Selected<%End If%>>V ��q</option>

                    <option value="CS" <% If ProductType = "CS" Then%>Selected<%End If%>>Select</option>

                    </select>        

        </td>

     </tr>

    

    <tr>

	    <td width="23%">§�鸹�X 

        </td> 

		<td width="37%">

					
<input name="Barcode" type="text" autocomplete="off" value="" size="20" maxlength="15" id="bc" ">

<input name="StationID" type="hidden" value="<%=StationID%>">

	<input type="Submit" Name="Button" value="�T�w">


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

</body>
</HTML>
