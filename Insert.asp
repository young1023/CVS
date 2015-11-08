 <!--#include file="include/SQLConn.inc" -->
<%


   Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
   StationID = Session("StationID")

   UserID    = Session("UserID")

   SecLevel  = Session("SecLevel")

   Barcode    = Request("Barcode")

   ProductType = Request("ProductType")

   CarPlate   = Request("CarPlate")


   Message = "ÅçÃÒ¦¨¥\!"

 
   Rs1.Open ("Exec InsertCoupon '"&Barcode&"', '"&StationID&"','"&ProductType&"','"&CarPlate&"','"&UserID&"' "), Conn, 3, 1


        


          Session("SecLevel") = SecLevel

          Session("UserID")   = UserID

          Session("StationID")= StationID

          Response.Redirect  "CouponVerification.asp?Message="&Message&"&CarPlate="&CarPlate&"&ProductType="&ProductType
       
%>
     