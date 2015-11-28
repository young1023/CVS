<!--#include file="include/SQLConn.inc" -->
<%

   UserID    = Request("UserID")

   Password  = Request("Password")

   StationID = Request("StationID")

   Set Rs1 = Server.CreateObject("Adodb.Recordset")

   Rs1.open ("Exec CheckLoginPwd '"&UserID&"', '"&Password&"' , '"&StationID&"' ") ,  conn,3,1

    If Not Rs1.EoF Then


   Response.Redirect "pco1.asp"


   Else

   Response.Redirect "Default.asp?Message=Fail"

   
   End if

%>
