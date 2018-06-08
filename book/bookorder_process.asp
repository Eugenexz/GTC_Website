<%
Session.timeout = 15
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "book","book","book"
Set Session("MyDB_conn") = conn
%>


<HTML>

<HEAD>
<title>Book Order</title>
</HEAD>

<BODY>



<%


strSQLQuery = "INSERT INTO Book (realname, CompanyName, MailingAddress, City, ProvState, Country, PostZipCode, Telephone, Fax, email, URL, BookType, Quantity, CardType, CardName, CardNumber, CardExpiry, OrderDate) VALUES('"& Request("realname") &"', '"& Request("CompanyName") &"','"& Request("MailingAddress") &"','"& Request("City") &"','"& Request("ProvState") &"','"& Request("Country") &"','"& Request("PostZipCode") &"','"& Request("Telephone") &"','"& Request("Fax") &"','"& Request("email") &"','"& Request("URL") &"','"& Request("BookType") &"','"& Request("Quantity") &"','"& Request("CardType") &"','"& Request("CardName") &"','"& Request("CardNumber") &"','"& Request("CardExpiry") &"','" & Now() &"')"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open strSQLQuery, conn, 3, 3


%>


<center>
<p>Your order was successfully received. You will be contacted regarding the status of your order.</p>
</center>


</BODY>

</HTML>


<%
conn.close
set conn = Nothing
%>
