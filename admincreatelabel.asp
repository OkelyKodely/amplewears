<%
dim xmlRequest1, orderid, rs

Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

orderid = request("orderid")

sqlstr = "select s_country, s_address, username, s_city, s_state, s_zipcode, shippingmethod from ""order2"" where orderid = " & orderid

set rs = Conn.execute(sqlstr)

dim address, username, city, state, zip, country, shippingmethod, weight

if not rs.eof then

  address = rs("s_address")
  username = rs("username")
  city = rs("s_city")
  state = rs("s_state")
  zip = rs("s_zipcode")
  'country = rs("s_country")
  country = "US"
  shippingmethod = rs("shippingmethod")

  sqlstr = "select a.*, b.weight from orderitems2 a inner join products2 b on a.itemid = b.itemid where a.orderid = " & orderid

  set rss = Conn.execute(sqlstr)
  weight = 0
  while not rss.eof
    weight = weight + cdbl(rss("weight"))
    rss.movenext
  wend

end if

xmlRequest1="<?xml version=""1.0""?>"
xmlRequest1=xmlRequest1 & "<AccessRequest xml:lang=""en-US"">"
xmlRequest1=xmlRequest1 & "<AccessLicenseNumber>9D5A1C57FAB32998</AccessLicenseNumber>"
xmlRequest1=xmlRequest1 & "<UserId>dhcho428</UserId>"
xmlRequest1=xmlRequest1 & "<Password>Coppersink21</Password>"
xmlRequest1=xmlRequest1 & "</AccessRequest>"
xmlRequest1=xmlRequest1 & "<ShipmentConfirmRequest>"
xmlRequest1=xmlRequest1 & "<Request>"
xmlRequest1=xmlRequest1 & "<TransactionReference>"
xmlRequest1=xmlRequest1 & "<CustomerContext>Customer Comment</CustomerContext>"
xmlRequest1=xmlRequest1 & "<XpciVersion/>"
xmlRequest1=xmlRequest1 & "</TransactionReference>"
xmlRequest1=xmlRequest1 & "<RequestAction>ShipConfirm</RequestAction>"
xmlRequest1=xmlRequest1 & "<RequestOption>validate</RequestOption>"
xmlRequest1=xmlRequest1 & "</Request>"
xmlRequest1=xmlRequest1 & "<LabelSpecification>"
xmlRequest1=xmlRequest1 & "<LabelPrintMethod>"
xmlRequest1=xmlRequest1 & "<Code>GIF</Code>"
xmlRequest1=xmlRequest1 & "<Description>gif file</Description>"
xmlRequest1=xmlRequest1 & "</LabelPrintMethod>"
xmlRequest1=xmlRequest1 & "<HTTPUserAgent>Mozilla/4.5</HTTPUserAgent>"
xmlRequest1=xmlRequest1 & "<LabelImageFormat>"
xmlRequest1=xmlRequest1 & "<Code>GIF</Code>"
xmlRequest1=xmlRequest1 & "<Description>gif</Description>"
xmlRequest1=xmlRequest1 & "</LabelImageFormat>"
xmlRequest1=xmlRequest1 & "</LabelSpecification>"
xmlRequest1=xmlRequest1 & "<Shipment>"
xmlRequest1=xmlRequest1 & "<RateInformation>"
xmlRequest1=xmlRequest1 & "<NegotiatedRatesIndicator/>"
xmlRequest1=xmlRequest1 & "</RateInformation>"
xmlRequest1=xmlRequest1 & "<Description/>"
xmlRequest1=xmlRequest1 & "<Shipper>"
xmlRequest1=xmlRequest1 & "<Name>priceBay</Name>"
xmlRequest1=xmlRequest1 & "<PhoneNumber>2136086228</PhoneNumber>"
xmlRequest1=xmlRequest1 & "<ShipperNumber>745VE8</ShipperNumber>"
xmlRequest1=xmlRequest1 & "<TaxIdentificationNumber>123456789</TaxIdentificationNumber>"
xmlRequest1=xmlRequest1 & "<Address>"
xmlRequest1=xmlRequest1 & "<AddressLine1>3266 Connecticut St. #303</AddressLine1>"
xmlRequest1=xmlRequest1 & "<City>Los Angeles</City>"
xmlRequest1=xmlRequest1 & "<StateProvinceCode>CA</StateProvinceCode>"
xmlRequest1=xmlRequest1 & "<PostalCode>90006</PostalCode>"
xmlRequest1=xmlRequest1 & "<PostcodeExtendedLow></PostcodeExtendedLow>"
xmlRequest1=xmlRequest1 & "<CountryCode>US</CountryCode>"
xmlRequest1=xmlRequest1 & "</Address>"
xmlRequest1=xmlRequest1 & "</Shipper>"
xmlRequest1=xmlRequest1 & "<ShipTo>"
xmlRequest1=xmlRequest1 & "<CompanyName>"&username&"</CompanyName>"
xmlRequest1=xmlRequest1 & "<AttentionName>"&username&"</AttentionName>"
xmlRequest1=xmlRequest1 & "<PhoneNumber></PhoneNumber>"
xmlRequest1=xmlRequest1 & "<Address>"
xmlRequest1=xmlRequest1 & "<AddressLine1>"&address&"</AddressLine1>"
xmlRequest1=xmlRequest1 & "<City>"&city&"</City>"
xmlRequest1=xmlRequest1 & "<StateProvinceCode>"&state&"</StateProvinceCode>"
xmlRequest1=xmlRequest1 & "<PostalCode>"&zip&"</PostalCode>"
xmlRequest1=xmlRequest1 & "<CountryCode>"&country&"</CountryCode>"
xmlRequest1=xmlRequest1 & "</Address>"
xmlRequest1=xmlRequest1 & "</ShipTo>"
xmlRequest1=xmlRequest1 & "<ShipFrom>"
xmlRequest1=xmlRequest1 & "<CompanyName>priceBay</CompanyName>"
xmlRequest1=xmlRequest1 & "<AttentionName>priceBay</AttentionName>"
xmlRequest1=xmlRequest1 & "<PhoneNumber>2136086228</PhoneNumber>"
xmlRequest1=xmlRequest1 & "<TaxIdentificationNumber>1234567877</TaxIdentificationNumber>"
xmlRequest1=xmlRequest1 & "<Address>"
xmlRequest1=xmlRequest1 & "<AddressLine1>3266 Connecticut St. #303</AddressLine1>"
xmlRequest1=xmlRequest1 & "<City>Los Angeles</City>"
xmlRequest1=xmlRequest1 & "<StateProvinceCode>CA</StateProvinceCode>"
xmlRequest1=xmlRequest1 & "<PostalCode>90006</PostalCode>"
xmlRequest1=xmlRequest1 & "<CountryCode>US</CountryCode>"
xmlRequest1=xmlRequest1 & "</Address>"
xmlRequest1=xmlRequest1 & "</ShipFrom>"
xmlRequest1=xmlRequest1 & "<PaymentInformation>"
xmlRequest1=xmlRequest1 & "<Prepaid>"
xmlRequest1=xmlRequest1 & "<BillShipper>"
xmlRequest1=xmlRequest1 & "<AccountNumber>745VE8</AccountNumber>"
xmlRequest1=xmlRequest1 & "</BillShipper>"
xmlRequest1=xmlRequest1 & "</Prepaid>"
xmlRequest1=xmlRequest1 & "</PaymentInformation>"
xmlRequest1=xmlRequest1 & "<Service>"
xmlRequest1=xmlRequest1 & "<Code>02</Code>"
xmlRequest1=xmlRequest1 & "<Description>"&shippingmethod&"</Description>"
xmlRequest1=xmlRequest1 & "</Service>"
xmlRequest1=xmlRequest1 & "<Package>"
xmlRequest1=xmlRequest1 & "<PackagingType>"
xmlRequest1=xmlRequest1 & "<Code>02</Code>"
xmlRequest1=xmlRequest1 & "<Description>Customer Supplied</Description>"
xmlRequest1=xmlRequest1 & "</PackagingType>"
xmlRequest1=xmlRequest1 & "<Description>Package Description</Description>"
xmlRequest1=xmlRequest1 & "<ReferenceNumber>"
xmlRequest1=xmlRequest1 & "<Code>00</Code>"
xmlRequest1=xmlRequest1 & "<Value>Package</Value>"
xmlRequest1=xmlRequest1 & "</ReferenceNumber>"
xmlRequest1=xmlRequest1 & "<PackageWeight>"
xmlRequest1=xmlRequest1 & "<UnitOfMeasurement/>"
xmlRequest1=xmlRequest1 & "<Weight>"&weight&"</Weight>"
xmlRequest1=xmlRequest1 & "</PackageWeight>"
xmlRequest1=xmlRequest1 & "<LargePackageIndicator/>"
xmlRequest1=xmlRequest1 & "<AdditionalHandling>0</AdditionalHandling>"
xmlRequest1=xmlRequest1 & "</Package>"
xmlRequest1=xmlRequest1 & "</Shipment>"
xmlRequest1=xmlRequest1 & "</ShipmentConfirmRequest>"

  Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  'xmlhttp.Open "POST","https://wwwcie.ups.com/ups.app/xml/ShipConfirm", false
  xmlhttp.Open "POST","https://onlinetools.ups.com/ups.app/xml/ShipConfirm", false 'production url endpoint
  xmlhttp.send xmlRequest1

  sResponseXML = xmlhttp.responseText

  a = instr(sResponseXML,"<ShipmentDigest>")
  b = instr(sResponseXML,"</ShipmentDigest>")

  sDisplayString = mid(sResponseXML, a+16, b-a-16)

dim xmlRequest2
xmlRequest2 = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
xmlRequest2=xmlRequest2 & "<AccessRequest>"
xmlRequest2=xmlRequest2 & "<AccessLicenseNumber>9D5A1C57FAB32998</AccessLicenseNumber>"
xmlRequest2=xmlRequest2 & "<UserId>dhcho428</UserId>"
xmlRequest2=xmlRequest2 & "<Password>Coppersink21</Password>"
xmlRequest2=xmlRequest2 & "</AccessRequest>"
xmlRequest2=xmlRequest2 & "<?xml version=”1.0″ encoding=”ISO-8859-1″?>"
xmlRequest2=xmlRequest2 & "<ShipmentAcceptRequest>"
xmlRequest2=xmlRequest2 & "<Request>"
xmlRequest2=xmlRequest2 & "<TransactionReference>"
xmlRequest2=xmlRequest2 & "<CustomerContext>Customer Comment</CustomerContext>"
xmlRequest2=xmlRequest2 & "</TransactionReference>"
xmlRequest2=xmlRequest2 & "<RequestAction>ShipAccept</RequestAction>"
xmlRequest2=xmlRequest2 & "<RequestOption>1</RequestOption>"
xmlRequest2=xmlRequest2 & "</Request>"
xmlRequest2=xmlRequest2 & "<ShipmentDigest>"&sDisplayString&"</ShipmentDigest>"
xmlRequest2=xmlRequest2 & "</ShipmentAcceptRequest>"

  Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  'xmlhttp.Open "POST","https://wwwcie.ups.com/ups.app/xml/ShipAccept", false
  xmlhttp.Open "POST","https://onlinetools.ups.com/ups.app/xml/ShipAccept", false 'production url endpoint
  xmlhttp.send xmlRequest2

  sResponseXML = xmlhttp.responseText

  a = instr(sResponseXML,"<GraphicImage>")
  b = instr(sResponseXML,"</GraphicImage>")

  sDisplayString = mid(sResponseXML, a+14, b-a-14)

  dim trackingNo

  a = instr(sResponseXML,"<TrackingNumber>")
  b = instr(sResponseXML,"</TrackingNumber>")

  trackingNo = mid(sResponseXML, a+16, b-a-16)

  Conn.close()
%>
<img src="data:image/gif;base64, <%=sDisplayString%>" />