<!--#include file="include/SQLConn.inc" -->

<%

UserIPAddress = Request.ServerVariables("REMOTE_ADDR")

Set Rs1 = Server.CreateObject("Adodb.Recordset")
  
sql1 = "Select * From Station s Left Join CVSUser u on "

sql1 = sql1 & " s.LoginUser = u.UserName Where IPAddress ='" &UserIPAddress& "'"

Set Rs1 = Conn.Execute(sql1)

UserID        = trim(Rs1("LoginUser"))

StationID     = trim(Rs1("Station"))

SecLevel      = trim(Rs1("SecLevel"))

If UserID = "" Then

   Response.Redirect "Default.asp"

End If


Message       = Request("Message")

Coupon_Number = Request("Coupon_Number")

Car_No        = Request("Car_No")

Face_Value    = Request("Face_Value")

Period        = Request("Period")


' search by date


If Request("SDay") <> "" then

SDay          = Request("SDay")

SMonth        = Request("SMonth") 

SYear         = Request("SYear")


Else

SDay = Day(now())

SMonth = Month(now())

SYear = year(now())

End If

If Request("NDay") <> "" then

NDay          = Request("NDay")

NMonth        = Request("NMonth") 

NYear         = Request("NYear")


Else

NDay = Day(now())

NMonth = Month(now())

NYear = year(now())

End If

Search_Date = SDay & "/" & SMonth & "/" & SYear

Search_NDate = NDay & "/" & NMonth & "/" & NYear

'response.write UserID

%>

<HTML>
<HEAD>
<title>
    --  禮券驗證系統 (澳門) -- 
</title>


<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" /> 
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
</HEAD>
<body class="homepage" onload="document.fm1.Barcode.focus();">



<!--#include file="include/header.inc" -->



<form name="fm1" method="post" action="">

<div align="center">

<h1 class="Title">禮券驗證系統 (澳門)</h1>

<span class="noprint">
<table width="60%" border="0" class="Report">

      <tr>

              <td>

                      <a href="CouponVerification.asp">驗證</a>                 

              </td>

               <td>

                      <a href="logout.asp">登出</a>                 

              </td>

      </tr>

</table>

<br/>
</span>
 <table width="98%" border="1" cellpadding="04" cellspacing="2" class="Ver">
                      
                        


         
			  <tr>
							
                                <td height="28"> 
                                  <%

		pageid=trim(request.form("pageid"))

		if pageid="" then
		  pageid=1
		end if

        Barcode = replace(trim(request.form("Barcode")),"%","％")

        Barcode = replace(Barcode,"'","''")

        If len(Barcode) = 15 then

        Face_Value = Left(Barcode, 3)

        Coupon_Type = Mid(Barcode, 5,2)

        Coupon_batch = Mid(Barcode,7,3)

        Coupon_Number = Mid(Barcode, 10,6)

        end if
        

' Start the Queries
    ' Start the queries
' *****************
      
       fsql = "select * from MasterCoupon where RequestedID = "&StationID

  ' Search Coupon Number
  ' ********************
   if Barcode <> "" then

       If len(Barcode) = 15 then

       fsql = fsql & " and Coupon_Number = '"& Coupon_Number &"'"

       fsql = fsql & " and Face_Value = '"& Face_Value &"'"

       fsql = fsql & " and Coupon_Type = '"& Coupon_Type &"'"

       fsql = fsql & " and Coupon_Batch = '"& Coupon_Batch &"'"

       Else

       fsql = fsql & " and Coupon_Number LIKE '%"&Barcode&"%' " 

       End If

   end if


 
  ' Search by Date
  ' **************


     fsql = fsql & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

  
      fsql = fsql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 


   ' By UserID

    if Trim(SecLevel) = 1  then


      fsql = fsql & " and Period = '"& UserID &"' " 

    end if
   
 
       fsql = fsql & " order by id desc"

        'response.write fsql
        'response.end

        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

%>

<span class="noprint">

<%

        if frs.RecordCount=0 then
           response.write "<font color=red>No Record</font>"
           'response.end
        else


          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ;"

  
         frs.PageSize = 100

         call countpage(frs.PageCount,pageid)

         end if

%>


	     &nbsp;&nbsp;<input type="text" name="Barcode" size="15" maxlength=15 value="">
		 &nbsp;&nbsp;<input type="button" value="禮券編號" onClick="findenum();">

</span>	   


                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 


   <table border="0" align=center cellpadding="4" cellspacing="1" class="DailyReport" width="100%" height="100%">

<tr bgcolor="#DFDFDF">
          <td colspan="7" align="right">油站</td>
          <td align="center"><% = StationID %></td>
      </tr>
    <tr> 
   

			<td colspan="3"> 日期:
			 
			<select name="SDay" class="common" class="reportdate">

        <% 
              j = 1

            for i = 1 to 31 

         %>

	<option value="<%=j%>" <% if trim(SDay) = trim(j) Then response.write "selected" %>><% = j %></option>
		
     <%
     
             j = j + 1

             Next 

         %>
		
			
			</select>


			<select name="SMonth" class="common"> 
 
        <% 
              k = 1

            for l = 1 to 12 

         %>
          	
    <option value="<%=k%>" <% if trim(SMonth) = trim(k) then response.write "selected"%>><% = Left(MonthName(k),3) %></option>
			 <%
     
             k = k + 1

             Next 

         %>

</select>


			<select name="SYear" class="common">   
<% 
Dim Year_starting
Dim Year_ending

Year_starting = Year(DateAdd("yyyy", -1, Now()))
year_ending = Year(Now())

for m = Year_starting to Year_ending
%>			         
			<option value="<%=m%>" <% if clng(m)=clng(SYear) then response.write "selected"%>><%=m%></option>

<% next %>

			</select> 
<span class="reportdate">

<% = SDay & " " & Left(MonthName(SMonth),3) & " " & SYear %>

</div>
			</td>


	<td colspan="3">
			 
			<select name="NDay" class="common">

        <% 
              j = 1

            for i = 1 to 31 

         %>

	<option value="<%=j%>" <% if trim(NDay) = trim(j) Then response.write "selected" %>><% = j %></option>
		
     <%
     
             j = j + 1

             Next 

         %>
		
			
			</select>


			<select name="NMonth" class="common"> 
 
        <% 
              k = 1

            for l = 1 to 12 

         %>
          	
    <option value="<%=k%>" <% if trim(NMonth) = trim(k) then response.write "selected"%>><% = Left(MonthName(k),3) %></option>
			 <%
     
             k = k + 1

             Next 

         %>

</select>


			<select name="NYear" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -1, Now()))

for i = Year_starting to Year(Now())
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(NYear) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> 

<span class="reportdate">

<% = NDay & " " & Left(MonthName(NMonth),3) & " " & NYear %>

</div>
			</td>

<td>更期:

            <select size="1" name="Shift" class="common">
            <option value="All">全日</option>
			<option value="11">早更</option>
			<option value="12">中更</option>
            <option value="13">晚更</option>
	</select>

</td>

	<td width="21%" ><input type="button" value="更結報告" onClick="Report();" class="noborder">
	</td>
    
</tr>


<% If i = 25 then %>
<tr bgcolor="#DFDFDF" class="page-break">
<% Else%>
<tr bgcolor="#DFDFDF">
<%End if%>
   

<td width="15%">日期 / 時間</td>
<td width="10%">更期</td>
<td width="15%">禮券編號</td>
<td width="10%">類型</td>
<td width="10%">產品</td>
<td width="10%">升數</td>
<td>銀碼</td>
<td>車牌</td>
</tr>

<%

Total = 0

i=0

 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
   if flage then
     mycolor="#ffffff"
   else
	 mycolor="#efefef"
   end if
  
%>


   <tr>

<% id = frs("id") %>
<td width="30%">
<% = i & ". " & frs("Present_Date") %>
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
<% = frs("Coupon_Number") %>
</td>
<td>
 <% = frs("Coupon_type") %>
</td>
<td>
<% = frs("Product_Type") %>
<td>
<% = FormatNumber(frs("SaleLitre"),2) %>
</td>
</td>
<td>

<%= frs("SaleAmount") %>
<%  

Total = Total + frs("SaleAmount") 

%>
   
</td>
<td>
<% = frs("Car_ID") %>
</td>

</tr>



<%

   flage=not flage
   frs.movenext
  loop
 end if
  %>
<tr>
<td colspan = "6" align="right">
Total:
</td>
<td colspan= "2" align="left"><% = Total %>
</td>
</tr>

                                  </table>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
<script language=JavaScript>
<!--

function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

function Report()
{
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

//-->
</script>

<span class="noprint">
<%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 response.write " First "
                 response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 response.write " Next "
                 response.write " Last "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
</span>
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center">
<span class="noprint"> 
<input type="button" value="   打印   "  onClick="window.print();" class='common'>
<% if SecLevel > 1 then %>
<input type="button" value="   CSV   "  onClick="window.doConvert();" class='common'>
<% end If %>
</span>

<%

              response.write "<input type=hidden value="&pageid&" name=pageid>"
			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                </td>
                              </tr>
                            </table>
                          </form>

                          


<%

  ' function
  Sub countpage(PageCount,pageid)
  response.write pagecount&"</font> Pages "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub


%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertReport.asp?StationID=<%=StationID%>&Level=<%=SecLevel%>&Sday=<%=SDay%>&SMonth=<%=SMonth%>&SYear=<%=SYear%>&NDay=<%=NDay%>&NMonth=<%=NMonth%>&NYear=<%=NYear%>&UserID=<%=UserID%>"); 

}

//-->
</SCRIPT>
</body>
</HTML>
