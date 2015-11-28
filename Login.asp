<!--#include file="include/SQLConn.inc" -->
<%

   UserID    = Request("UserID")

   Password  = Request("Password")

   StationID = Request("StationID")

   Dim UserIPAddress
   UserIPAddress = Request.ServerVariables("REMOTE_ADDR")


   Set Rs1 = Server.CreateObject("Adodb.Recordset")

   Rs1.open ("Exec CheckLoginPwd '"&UserID&"', '"&Password&"' , '"&UserIPAddress&"' ") ,  conn,3,1

   If Not Rs1.EoF Then

   Response.Redirect "CouponVerification.asp"

   Else

   Response.Redirect "Default.asp?Message=Fail"

   
   End if

%>
