<%
Option Explicit

' By default, this sample code is designed to post to our test server for
' developer accounts: https://test.authorize.net/gateway/transact.dll
' for real accounts (even in test mode), please make sure that you are
' posting to: https://secure.authorize.net/gateway/transact.dll
Dim post_url
post_url = "https://test.authorize.net/gateway/transact.dll"

Dim post_values
Set post_values = CreateObject("Scripting.Dictionary")
post_values.CompareMode = vbTextCompare

'the API Login ID and Transaction Key must be replaced with valid values

dim mailtothisfool

dim username

dim cartid

dim oConnection

dim s_address, s_city, s_state, s_zipcode, s_country, x_login, x_tran_key, x_card_num, x_exp_date, Amount, shipping

x_login = request.form("x_login")

x_tran_key = request.form("x_tran_key")

x_card_num = request.form("x_card_num")

x_exp_date = request.form("x_exp_date")

Amount = request.form("x_amount")

shipping = request.form("shipping")

username = request.form("x_custom_1")

cartid = request.form("x_custom_2")

s_address = request.form("x_custom_3")

s_city = request.form("x_custom_4")

s_state = request.form("x_custom_5")

s_zipcode = request.form("x_custom_6")

s_country = request.form("x_custom_7")

dim xmlRequest1

xmlRequest1="<?xml version=""1.0""?>"
xmlRequest1=xmlRequest1 & "<AccessRequest xml:lang=""en-US"">"
xmlRequest1=xmlRequest1 & "<AccessLicenseNumber>9D5A1C57FAB32998</AccessLicenseNumber>"
xmlRequest1=xmlRequest1 & "<UserId>dhcho428</UserId>"
xmlRequest1=xmlRequest1 & "<Password>coppersink21</Password>"
xmlRequest1=xmlRequest1 & "</AccessRequest>"
xmlRequest1=xmlRequest1 & "<?xml version=""1.0""?>"
xmlRequest1=xmlRequest1 & "<AddressValidationRequest xml:lang=""en-US"">"
xmlRequest1=xmlRequest1 & "<Request>"
xmlRequest1=xmlRequest1 & "<TransactionReference>"
xmlRequest1=xmlRequest1 & "<CustomerContext>Summary Description</CustomerContext>"
xmlRequest1=xmlRequest1 & "<XpciVersion>1.0</XpciVersion>"
xmlRequest1=xmlRequest1 & "</TransactionReference>"
xmlRequest1=xmlRequest1 & "<RequestAction>XAV</RequestAction>"
xmlRequest1=xmlRequest1 & "<RequestOption>3</RequestOption>"
xmlRequest1=xmlRequest1 & "</Request>"
xmlRequest1=xmlRequest1 & "<AddressKeyFormat>"
xmlRequest1=xmlRequest1 & "<AddressLine>"&s_address&"</AddressLine>"
xmlRequest1=xmlRequest1 & "<PoliticalDivision2>"&s_city&"</PoliticalDivision2>"
xmlRequest1=xmlRequest1 & "<PoliticalDivision1>"&s_state&"</PoliticalDivision1>"
xmlRequest1=xmlRequest1 & "<PostcodePrimaryLow>"&s_zipcode&"</PostcodePrimaryLow>"
xmlRequest1=xmlRequest1 & "<CountryCode>"&s_country&"</CountryCode>"
xmlRequest1=xmlRequest1 & "</AddressKeyFormat>"
xmlRequest1=xmlRequest1 & "</AddressValidationRequest>"

Dim oRequest, postresponse
Set oRequest = Server.CreateObject("Microsoft.XMLHTTP")
oRequest.open "POST", "https://onlinetools.ups.com/ups.app/xml/XAV", false
oRequest.send xmlRequest1
postresponse = oRequest.responseText
Set oRequest = nothing

dim s_address2

s_address2 = request.form("x_custom_8")

'''response.write postresponse

if instr(postresponse, "<AddressLine>") > 0 and len(s_address) > 0 then

post_values.Add "x_login", x_login
post_values.Add "x_tran_key", x_tran_key

post_values.Add "x_delim_data", "TRUE"
post_values.Add "x_delim_char", ","
post_values.Add "x_relay_response", "FALSE"

post_values.Add "x_type", "AUTH_CAPTURE"
post_values.Add "x_method", "CC"
post_values.Add "x_card_num", x_card_num
post_values.Add "x_exp_date", x_exp_date

dim sh
sh = split(shipping, "$")

Amount = cdbl(Amount) + cdbl(sh(1))
Amount = formatnumber(Amount,2)

post_values.Add "x_amount", Amount
post_values.Add "x_description", "Auth Cap Transaction"

' Additional fields can be added here as outlined in the AIM integration
' guide at: http://developer.authorize.net

' This section takes the input fields and converts them to the proper format
' for an http post. For example: "x_login=username&x_tran_key=a1B2c3D4"
Dim post_string
post_string = ""
dim a
For Each a In post_values
post_string=post_string & a & "=" & Server.URLEncode(post_values(a)) & "&"
Next
post_string = Left(post_string,Len(post_string)-1)

' The following section provides an example of how to add line item details to
' the post string. Because line items may consist of multiple values with the
' same key/name, they cannot be simply added into the above array.
'
' This section is commented out by default.
'
'Dim line_items(3)
'line_items(0) = "item1<|>golf balls<|><|>2<|>18.95<|>Y"
'line_items(1) = "item2<|>golf bag<|>Wilson golf carry bag, red<|>1<|>39.99<|>Y"
'line_items(2) = "item3<|>book<|>Golf for Dummies<|>1<|>21.99<|>Y"
'
'For Each item In line_items
' post_string = post_string & "&x_line_item=" & Server.URLEncode(item)
'Next

' We use xmlHTTP to submit the input values and record the response
Dim objRequest, post_response
Set objRequest = Server.CreateObject("Microsoft.XMLHTTP")
'objRequest.setRequestHeader "Content-length", post_string.length
objRequest.open "POST", post_url, false
objRequest.send post_string
post_response = objRequest.responseText
Set objRequest = nothing

' the response string is broken into an array using the specified delimiting character
Dim response_array
response_array = split(post_response, post_values("x_delim_char"), -1)

' the results are output to the screen in the form of an html numbered list.
dim continu
continu = 0

dim tid

dim cnt
cnt = 0
dim value
For Each value in response_array
  if cnt = 0 then
    if value = "1" then
      continu = 1
    end if
  elseif cnt = 6 then
    tid = value
  end if
  response.write value & ","
  cnt = cnt + 1
Next
' individual elements of the array could be accessed to read certain response
' fields. For example, response_array(0) would return the Response Code,
' response_array(2) would return the Response Reason Code.
' for a list of response fields, please review the AIM Implementation Guide


'response.write continu


'response.end

if continu <> 1 then
'response.write Amount
  response.redirect "paymentfailed.asp"
else

dim oRS

'set oConnection = server.createobject("ADODB.Connection")

'oConnection.Open "odbc2","sa","coppersink21"

dim Conn

Set Conn = Server.Createobject("ADODB.Connection")

Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

dim sqlstr

sqlstr = "SELECT * FROM shoppers2 WHERE username='" & username & "'"
'response.write sqlstr
set oRS = Conn.Execute(sqlstr)

if not oRS.eof then
  mailtothisfool = oRS("email")
end if

Dim ShippingMethod, ShippingCost
ShippingMethod = replace(sh(0), " - ", "")
ShippingCost = formatnumber(sh(1), 2)
sqlstr = "INSERT INTO ""order2"" (transid, cc4, status, shippingmethod, shippingcost, username, price, inputdate, s_address, s_city, s_state, s_zipcode, s_country) VALUES "
sqlstr = sqlstr & "('"&tid&"', '"&right(x_card_num, 4)&"', 'processing...', '"&ShippingMethod&"',"&ShippingCost&",'"&username&"',"&Amount&",sysdate,'"&s_address&" "&s_address2&"','"&s_city&"','"&s_state&"','"&s_zipcode&"','"&s_country&"')"
Conn.Execute(sqlstr)

dim rs

sqlstr = "SELECT orderid FROM ""order2"" ORDER BY orderid DESC"
set rs = Conn.Execute(sqlstr)

dim orderid

if not rs.eof then

  orderid = rs("orderid")

  sqlstr = "SELECT * FROM cart2 where cid = " & cartid

  dim r_s

  set r_s = Conn.Execute(sqlstr)

  while not r_s.eof

    dim itemid

    itemid = r_s("itemid")

    sqlstr = "INSERT INTO orderitems2 (orderid,itemid) VALUES ("&orderid&","&itemid&")"
    'response.write sqlstr
    Conn.Execute(sqlstr)

    r_s.movenext

  wend

end if

sqlstr = "DELETE FROM cart2 WHERE cid = " & cartid
'response.write sqlstr
Conn.Execute(sqlstr)

Conn.close()

dim objMail 
Set objMail = Server.CreateObject("CDO.Message") 

dim smtpServer, yourEmail, yourPassword
smtpServer = "smtp.gmail.com"
yourEmail = "sale.pbay@gmail.com"     'replace with a valid gmail account
yourPassword = "coppersink21"   'replace with a valid password for account set in yourEmail 

objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = yourEmail
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = yourPassword
objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
objMail.Configuration.Fields.Update 

objMail.Subject="Your AmpleWears Order"
objMail.htmlBody = "You have bought some items from AmpleWears.ml in the amount of $" & Amount & " USD.  Thank you for doing business with us.  We will process your order shortly.<br><br>priceBay team"
objMail.From = "sale.pbay@gmail.com"
objMail.To = mailtothisfool
objMail.Send   

response.redirect "thankyou.asp"

end if
  
else
%>
<script type="text/javascript">window.location = document.referrer + '&index=1';</script>
<%
end if
%>