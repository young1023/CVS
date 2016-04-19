<%
Set Conn = Server.CreateObject("ADODB.Connection")
'Strnn = "Driver=CVS;Server=10.122.154.9;Database=HKMO_CVS_DB;Trusted_Connection=yes;"
StrCnn = "Provider=sqloledb; User ID=intranet; Password=intranet; Initial Catalog=CVS; Data Source=localhost"
'StrCnn = "Provider=sqloledb; User ID=intranet; Password=intranet; Initial Catalog=CVS; Data Source=localhost"
'StrCnn = "Driver={SQL Server};Server=10.122.154.9\SS05;Database=HKMO_CVS_DB;Trust_Connection=yes;"
Conn.CommandTimeout=0
Conn.ConnectionTimeout=0
Conn.Open StrCnn

%>
