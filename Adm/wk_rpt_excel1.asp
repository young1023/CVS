<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=Weekly_Rpt_"&Month(Now())&Year(now())&".xls"

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




'**************
'Initialisation
'**************
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001
	

' Create a server recordset object

       set rs = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") 
	  rs.open ("Exec WeeklyReport '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") ,  conn,3,1

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

<div align="center">

<table BORDER="1" width="98%">

     <tr bgcolor="#DFDFDF">

<td height="28">From Date</td>
<td height="28">To Date</td>
<td height="28">(Sum) Total / Amount (Sum)</td>
<td  height="28">Print Excel</td>
<td  height="28">Excel Type</td>
<td  height="28">Station</td>
<td >Ship To Code</td>
<td  height="28">Sold to Code</td>
</tr>
  

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

   <tr>

<td align=center width="95" height="28"><% = rs("Doc Date")%></td>
<td  height="28"><% = rs("To Date") %>
</td>

<td height="28"><% = rs("Amount") %>
</td>

<td  height="28"><% = rs("Print Excel") %>
</td>

<td  height="28">
<% = rs("Excel_Type") %>
</td>

<td >
<% = rs("ShipToCode") %>
</td>

<td>
<% = rs("SoldToCode") %> 
</td>

</tr>

<%
' Move to the next record
rs.movenext
' Loop back to the do statement
loop %>
</table>

</div>

</body>
</html>

<%
' Close and set the recordset to nothing
rs.close
set rs=nothing
%>