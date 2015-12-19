<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=Redemp_Rpt_"&Month(Now())&Year(now())&".xls"

%>

<!--#include file="include/SQLConn.inc" -->

<%


'**************
'Argument handler
'**************

	
From_Date   = Request.Form("From_Date")
To_Date     = Request.Form("To_Date")


'**************
'Initialisation
'**************
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001
	

' Create a server recordset object

      set rs = server.createobject("adodb.recordset")
      'response.write  ("Exec RedemptionReport1 '"&From_Date&"', '"&To_Date&"' ") 
	  rs.open ("Exec RedemptionReport1 '"&From_Date&"', '"&To_Date&"' ") ,  conn,3,1

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
<tr>
<td height="28">Present Date</td>
<td height="28">Station</td>
<td height="28">Total</td>

</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr>

<td align=center width="95" height="28"><% = rs("Present Date")%></td>
<td  height="28"><% = rs("Station") %>
</td>

<td height="28"><% = rs("Total") %>
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