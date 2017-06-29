<!--#include file="include/SQLConn.inc" -->
<%

   UserID    = Request("UserID")

   Password  = Request("Password")

  
   Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
   sql1 = "Select * From CVSUser Where UserName ='" &UserID& "' and Password ='" &Password& "'"

   Set Rs1 = Conn.Execute(sql1)

  
   If Not Rs1.EoF Then

   Session("SecLevel") = Rs1("SecLevel")

   Response.Redirect "Status1.asp"


   Else

   Response.Redirect "Default.asp?Message=Fail"

   
   End if

%>
