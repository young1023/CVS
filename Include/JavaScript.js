<SCRIPT language=JavaScript>
<!--

function doLogin(){

   document.fm1.action="Login.asp";
   
   

   
   if (document.fm1.UserID.value == "")
{
   alert("�п�J�Τ�s��.");
   document.fm1.UserID.focus();
   return false;
   }

 if (document.fm1.Password.value == "")
  
{
   alert("�п�J�K�X.");
   document.fm1.Password.focus();
   return false;
   }	
   document.fm1.submit();
}



// Check digit
function mod10CheckDigit() {
//Get barCode field value, Strip all characters
//except numbers Repopulate field w/ new value
bc = document.barCodeForm.Barcode.value
barcode1 = document.barCodeForm.Barcode.value
checkdigit = bc.substr(3,1);
bc = bc.replace(/[^0-9]+/g,'');
document.barCodeForm.Barcode.value = bc;
total = 0;
//Get Even Numbers
for (i=bc.length-2; i>=0; i=i-2) {
if (i != 3){
tempeven = parseInt(bc.substr(i,1)) ;
if (tempeven > 9) {
tempeven = tempeven - 9;
}
total = total + tempeven ;
}
}
//Get Odd Numbers
for (i=bc.length-1; i>=0; i=i-2) {
temp = parseInt(bc.substr(i,1)) * 2 ;
if (temp > 9) {
temp = temp - 9;
}
total = total + temp;
}

//Determine the checksum. if modDigit = 10, modDigit = 0
modDigit = (10 - ((total + 17) % 10)) % 10;

var CouponType = bc.substr(4,2);
var ProductType = document.getElementById("ProductType");

if (CouponType == 73 || CouponType == 83){
if (ProductType.options[ProductType.selectedIndex].value === "CS"){
alert("���ҥ��� - ��������!");
return false;
}
}

if (CouponType == 24){
if (ProductType.options[ProductType.selectedIndex].value != 54){
alert("���ҥ��� - ��������!");
return false;
}
}

if (checkdigit != modDigit){
alert("���ҥ��� - §��L��!");
return false;
}


//Populate the modDigit field with the value
//document.barCodeForm.barCodeShow.value = barcode1;
document.barCodeForm.action="execute.asp";
document.barCodeForm.submit();
//return false;
}


//-->
</SCRIPT>