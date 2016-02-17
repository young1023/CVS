<%
' Tells the browser to open excel
'Response.ContentType = "application/vnd.ms-excel" 
'Response.addHeader "content-disposition","attachment;filename=Weekly_Rpt_"&Month(Now())&Year(now())&".xls"

%>

<!--#include file="include/SQLConn.inc" -->

<%


'**************
'Argument handler
'**************

From_Date   = Request.Form("From_Date")
To_Date     = Request.Form("To_Date")
Print_Excel = Request.Form("Print_Excel")
Excel_Type  = Request.Form("Excel_Type")


     ' Restrieve Excel Type

      set rs0 = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport_Excel_Type '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"' ") 
	  rs0.open ("Exec WeeklyReport_Excel_Type '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"'") ,  conn,3,1

   

' Move to the first record
rs0.movefirst

' Start a loop that will end with the last record
do while not rs0.eof
 
	
' Create a server recordset object

      set rs = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&rs0("Excel_Type")&"' ") 
	  rs.open ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&rs0("Excel_Type")&"' ") ,  conn,3,1

%>


<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<body>
<Head>
<STYLE TYPE="text/css">
<!--

TD 
{
  color: black;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}

TD.caption
{
  color: red;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}
-->
</STYLE>
</head>


<table BORDER="1" width="98%">
<tr bgcolor="#DFDFDF">
<td height="28">From Date</td>
<td height="28">To Date</td>
<td height="28">Station</td>
<td height="28">Print Excel</td>
<td height="28">Excel Type</td>
<td height="28">Sold to Code</td>
<td height="28">&nbsp;</td>
<td height="28">(Sum) Total / Amount (Sum)</td>
<td height="28">&nbsp;</td>
<td height="28">&nbsp;</td>
<td >Ship To Code</td>
</tr>
  

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

   <tr>

<td height="28"><% = rs("Doc Date")%>
</td>

<td  height="28"><% = rs("To Date") %>
</td>

<td  height="28">
<% = rs("Station") %>
</td>

<td  height="28"><% = rs("Print Excel") %>
</td>

<td  height="28">
<% = rs("Excel_Type") %>
</td>

<td>
<% = rs("SoldToCode") %> 
</td>

<td></td>

<td height="28"><% = rs("Amount") %>
</td>

<td></td>

<td></td>

<td >
<% = rs("ShipToCode") %>
</td>



</tr>

<%
' Move to the next record
rs.movenext
' Loop back to the do statement
loop 

%>

</table>


<br>


<%

' Calculate the total

      set frs1 = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&rs0("Excel_Type")&"' ") 
	  frs1.open ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&rs0("Excel_Type")&"' ") ,  conn,3,1
%>


 <table border="0" cellpadding="1" width="40%" cellspacing="1" class="normal">

     <tr bgcolor="#DFDFDF">

<%

    ' Total
    do while not frs1.EoF
 
  %>

  <tr>

<td>Total Amount of Excel Type: <% = frs1("Excel_Type") %>
</td>

<td height="28"><% = frs1("Total") %>
</td>

  </tr>

<%
   
   frs1.movenext
  loop

%>


<%
' Move to the next record
rs0.movenext
' Loop back to the do statement
loop 

%>



</body>
</html>

<%
' Close and set the recordset to nothing
rs.close
set rs=nothing
frs1.close
set frs1=nothing
%>