<!--#include file="include/SQLConn.inc" -->
<%

StationID     = Request("Station")

' Tells the browser to open excel
'Response.ContentType = "application/vnd.ms-excel" 
'Response.addHeader "content-disposition","attachment;filename=DailyReport_Station"&StationID&"_"&Month(Now())&Year(now())&".xls"

' Received parameters 

SDay          = Request("SDay")

SMonth        = Request("SMonth") 

SYear         = Request("SYear")

NDay          = Request("NDay")

NMonth        = Request("NMonth") 

NYear         = Request("NYear")

UserID        = Request("UserID")

Search_Date = SDay & "/" & SMonth & "/" & SYear

Search_NDate = NDay & "/" & NMonth & "/" & NYear


%>

<HTML>
<HEAD>
<title>
    --  禮券驗證系統 -- 
</title>


<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" /> 
</HEAD>
<body>
<%

Total = 0
Total2 = 0
Total3 = 0

' Start the Queries
' *****************
      
   fsql = "SELECT m.Product_Type as ProductType, * FROM MasterCoupon m INNER JOIN CouponRequest c "

   fsql = fsql & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

   fsql = fsql & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

   fsql = fsql & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

 
  ' Search by Date
  ' **************


    fsql = fsql & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

  
    fsql = fsql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

    
    if UserID <> "All" then

    fsql = fsql & " and Period = '"& UserID &"' " 

    end if

    fsql = fsql & " order by id desc"

    'response.write fsql

    set frs= conn.execute(fsql)
	
%>


<table>
<tr bgcolor="#DFDFDF">
<td >日期 / 時間</td>
<td >更期</td>
<td >銀碼</td>
<td >類型</td>
<td >批次</td>
<td >禮券編號</td>
<td >產品</td>
<td >機號</td>
<td >電子禮券</td>
</tr>

<%

  do while not frs.eof

   if frs.eof then exit do
   
  
%>

<tr>
<td >
<% = frs("Present_Date") %>
</td>

<td>
<%
If frs("Period") = 11 Then

   Response.write "早"

Elseif frs("Period") = 12 Then

   Response.write "中"

Elseif frs("Period") = 13 Then

   Response.write "晚"

End if
%>
</td>

<td>
<%= frs("Face_Value") %>
<%  

Total = Total + frs("Face_Value") 

%>
</td>

<td>
 <% = frs("Coupon_type") %>
</td>

<td>
 <% = frs("Coupon_Batch") %>
</td>

<td>
<% = frs("Coupon_Number") %>
</td>

<td>
<% = frs("ProductType") %>
</td>


<td>
<% = frs("MachineNo") %>

</td>

<td>
<% = frs("Digital") %>
</td>

</tr>



<%

   
   frs.movenext

  loop
 

  %>

<tr>
<td colspan = "2" align="right">
Total:
</td>
<td colspan= "8" align="left"><% = Total %>
</td>
</tr>

<%

   ' Start the Queries for eCoupon  
   ' *****************************
      
     sql = "SELECT MachineNo, Digital, Face_Value, Count(Face_Value) as eCount, Sum(Cast(Face_Value as float)) as eAmount FROM MasterCoupon m INNER JOIN CouponRequest c "

     sql = sql & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

     sql = sql & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

     sql = sql & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

     sql = sql & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

     sql = sql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

   ' By UserID

    if UserID <> "All"  then

      sql = sql & " and Period = '"& UserID &"' " 

    end if 
 
      sql = sql & " Group by MachineNo, Digital, Face_Value"

      'response.write sql


     Set rs = Conn.Execute(sql)


  %>

     <tr bgcolor="#DFDFDF">
  
          <td colspan="5">(總結 (以銀碼分類))</td>

     </tr>

             <tr>

                    <td >電腦編號</td>

                    <td>電子禮券</td>

                    <td>銀碼</td>

                    <td>數量</td>

                    <td>總值</td>

             </tr>

<%

    If Not rs.EoF Then

       rs.MoveFirst

     Do While Not rs.Eof
 
  %>           
             <tr>

                    <td><% = rs("MachineNo") %></td>

                    <td><% = rs("Digital") %></td>

                    <td><% = rs("Face_Value") %></td>

                    <td><% = rs("eCount") %></td>

                    <td><% = rs("eAmount") %>

<%  

Total2 = Total2 + rs("eAmount") 

%>


                     </td>


             </tr>
<%
            rs.MoveNext

            Loop

   End If

%>

  <tr>
                    <td colspan="4"></td>
                    <td><% = Total2 %></td>

  </tr>   

<%

   ' Start the Queries for eCoupon  
   ' *****************************
      
     sql2 = "SELECT MachineNo, Digital, Coupon_type, Batch, Count(Batch) as eCount2, Sum(Cast(Face_Value as float)) as eAmount2 FROM MasterCoupon m INNER JOIN CouponRequest c "

     sql2 = sql2 & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

     sql2 = sql2 & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

     sql2 = sql2 & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

     sql2 = sql2 & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

     sql2 = sql2 & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

   ' By UserID

    if UserID <> "All"  then

      sql2 = sql2 & " and Period = '"& UserID &"' " 

    end if 
 
      sql2 = sql2 & " Group by MachineNo, Coupon_type, Digital, Batch, Face_Value"

      'response.write sql2


     Set rs2 = Conn.Execute(sql2)


  %>

    <tr bgcolor="#DFDFDF">
  
          <td colspan="6">總結 (以批次分類)</td>

     </tr>
          

             <tr>

                    <td >電腦編號</td>

                    <td>電子禮券</td>

                    <td >類型</td>

                    <td>批次</td>

                    <td>數量</td>

                    <td>總值</td>

             </tr>

<%

    If Not rs2.EoF Then

       rs2.MoveFirst

     Do While Not rs2.Eof
 
  %>           
             <tr>

                    <td><% = rs2("MachineNo") %></td>

                    <td><% = rs2("Digital") %></td>

                    <td ><% = rs2("Coupon_Type") %></td>

                    <td><% = rs2("Batch") %></td>

                    <td><% = rs2("eCount2") %></td>

                    <td><% = rs2("eAmount2") %>

<%  

Total3 = Total3 + rs2("eAmount2") 

%>

</td>


             </tr>
<%
            rs2.MoveNext

            Loop

   End If

%>
  <tr>
                    <td colspan="5"></td>
                    <td><% = Total3 %></td>

  </tr>   

         </table>
                                
</body>
</HTML>
