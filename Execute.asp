 <!--#include file="include/SQLConn.inc" -->
<%

   IPAddress = Request.ServerVariables("REMOTE_ADDR")

   IPAddress = "192.168.1.1"

   ScanDate = Now()



   Set Rs0 = Server.CreateObject("Adodb.Recordset")
   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")
   Set Rs3 = Server.CreateObject("Adodb.Recordset")
 
   Barcode    = trim(Request("Barcode"))

   ProductType = trim(Request("ProductType"))


    ' Check digit 
    If Barcode = "Invalid" Then

           Message = "驗證失敗 - 禮券號碼不正確, 請重試!  "


    Response.Redirect  "CouponVerification.asp?&ProductType="&ProductType&"&Color=2&Message="&Message
      
 
     End If
 
   
   Set Message = Nothing

   ' Check Coupon Type
  
   Rs0.open ("Exec Check_CouponType '"&Barcode&"', '"&ProductType&"'") ,  conn,3,1

   'Response.write ("Exec Check_CouponType '"&Barcode&"', '"&ProductType&"'")

   'Response.end

   If Rs0("Tcount") = 0 Then

           Message = "驗證失敗 - 產品類型不符!"


    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      
 
   End if

          Rs0.close
          set Rs0 = nothing



   ' Check Barcode Range

    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

    'Response.write ("Exec Checkrange '"&Barcode&"'")

   'Response.end

   If Rs1.EoF Then

           Message = "驗證失敗 - 禮券無效 "


    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      
 
   Else

       RequestID = Rs1("RequestID")


       If Rs1("Expiry_Date") < DateValue(Now) Then

   
          Message = "驗證失敗 - 禮券過期!"
        
          

    Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
      

       End if

   End if


          Rs1.close
          set Rs1 = nothing


      ' --------------------------------------------

      ' Enhancement of restricted hours and stations

      ' --------------------------------------------

      ' Check restricted time

          Set Rs4 = Server.CreateObject("Adodb.Recordset")

          Rs4.open ("Exec Check_Time '"& RequestID &"', '"& ScanDate &"' ") ,  conn,3,1

          'Response.write ("Exec Check_Time '"& RequestID &"', '"& ScanDate &"' ") 

         If Not Rs4.EoF Then


                      Message = "驗證失敗 - 禮券只可在 "& Rs4("Start_Time") &" 時至 "& Rs4("End_Time") &" 時使用!"


       Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message

     
           End If

          Rs4.close
          set Rs4 = nothing

       ' Check restricted station

          Set Rs5 = Server.CreateObject("Adodb.Recordset")

          Rs5.open ("Exec Check_Station '"& RequestID &"', '"& IPaddress &"' ") ,  conn,3,1

          'Response.write ("Exec Check_Station '"& RequestID &"', '"& IPaddress &"' ") 

          'Response.end

         If Rs5("Tcount") = 0 Then


                      Message = "驗證失敗 - 禮券只可在指定油店使用!"


          Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message

     
           End If

          Rs5.close
          set Rs5 = nothing

       ' --------------------------------------------

          Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

          Response.write ("Exec CheckCouponExist '"&Barcode&"'")

         ' If there is record
         If Not Rs2.EoF Then



 

               Present_Date = Rs2("Present_Date")

               RequestedID  = Rs2("RequestedID")

               StationName  = Rs2("StationName")

         'response.write Present_Date
         'response.end
              Color   = "2"

              Message = "驗證失敗 - 重複使用! 禮券 " & Barcode '& " 曾於 " & Present_Date & " 在"& StationName & "站驗證!"

              Message2 = " 曾於 " & Present_Date & " 在"& StationName & "站驗證!"

              Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color="&Color&"&Message="&Message&"&Message2="&Message2


 
         Else


  Response.write ("Exec InsertCoupon '"&Barcode&"', '"&IPAddress&"','"&ProductType&"','"&ScanDate&"'")

          Rs3.Open ("Exec InsertCoupon '"&Barcode&"', '"&IPAddress&"','"&ProductType&"','"&ScanDate&"'"), Conn, 3, 1

              ' If record is inserted.
              If Not Rs3.EoF Then 
             
                  Message = "驗證成功!  號碼: " &  Rs3("Coupon_number") 
  
                  Message2 = "  時間: " & Rs3("Present_Date")

                   Color   = "1"

                   Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color="&Color&"&Message="&Message&"&Message2="&Message2
      
                    Else

                  Message = "驗證失敗 - 請重試! "
               

                Response.Redirect  "CouponVerification.asp?ProductType="&ProductType&"&Color=2&Message="&Message
             


              End if
    
       
      

       
          End if


          Rs2.close
          set Rs2 = nothing

 
%>
     