<%
' Tells the browser to open excel
Response.ContentType = "text/csv"
Response.addHeader "content-disposition","attachment;filename=Weekly_Rpt_"&Month(Now())&Year(now())&".csv"

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

  

<%

Response.Write "Doc Date,To Date,Station,Print Excel,Excel Type,Sold To Code,,(Sum) Total / Amount (Sum),,,Ship To Code" & vbcrlf 

If not rs.eof then

' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
Response.Write """" & FormatDateTime(rs("Doc Date"),2) & """," 
Response.Write """" & FormatDateTime(rs("To Date"),2) & """," 
Response.Write """" & rs("Station") & """," 
Response.Write """" & rs("Print Excel") & """," 
Response.Write """" & rs("Excel_Type") & """," 
Response.Write """" & rs("SoldToCode") & """," 
Response.Write "," 
Response.Write """" & rs("Amount") & """," 
Response.Write "," 
Response.Write "," 
Response.Write """" & rs("ShipToCode") & """" & vbCrLf  

' Move to the next record
rs.movenext
' Loop back to the do statement
loop 

end if


' Calculate the total

      set frs1 = server.createobject("adodb.recordset")
      'response.write  ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") 
	  frs1.open ("Exec WeeklyReport_Total '"&From_Date&"', '"&To_Date&"', '"&Print_Excel&"', '"&Excel_Type&"' ") ,  conn,3,1


    ' Total
    do while not frs1.EoF


Response.Write "Total Amount of Excel Type:""" & frs1("Excel_Type") & """," 

Response.Write """" & frs1("Total") & """" & vbCrLf  
 

   
   frs1.movenext
  loop

%>




<%
' Close and set the recordset to nothing
rs.close
set rs=nothing
frs1.close
set frs1=nothing
%>