<!--#include file="include/SQLConn.inc" -->
<%

StationID     = Request("Station"))

' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=DailyReport_Station"&StationID&"_"&Month(Now())&Year(now())&".xls"

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

    
    if UserID <> "" then

    fsql = fsql & " and Period = '"& UserID &"' " 

    end if

    fsql = fsql & " order by id desc"

    set frs= conn.execute(fsql)
	
%>


<table>
<tr bgcolor="#DFDFDF">
<td width="15%">日期 / 時間</td>
<td width="10%">更期</td>
<td width="10%">銀碼</td>
<td width="10%">類型</td>
<td width="10%">批次</td>
<td width="15%">禮券編號</td>
<td width="10%">產品</td>
<td width="10%">機號</td>
<td width="10%">電子禮券</td>
</tr>

<%

  do while not frs.eof

   if frs.eof then exit do
   
  
%>

<tr>
<td width="30%">
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
<% = frs("Product_Type") %>
</td>

<td>

</td>

<td>
<% = frs("Digital") %>
</td>

</tr>



<%

   
   frs.movenext

  loop
 

  %>

</table>
                                
</body>
</HTML>
