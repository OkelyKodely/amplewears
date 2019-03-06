<!--#include file="inc/header.inc"-->
  <h1>&nbsp;</h1>
  <h1 style="position:relative;left:0px;top:0px">Checkout</h1>

  <a href="cart.asp" style="color:gray"><h3 style="position:relative;left:0px;top:0px;">View Your Cart</h3></a>

  <div class="shop">
  
    <div class="shoptitle" style="position:relative;top:0px">
      <span style="font-size:23">Please Enter Your Credit Card & Shipping Information</span><br><br>
    </div>

    <div class="shoptitle" style="position:relative;top:0px" id="steps">
      <img src="/graphics/images/step1.png"><br><br>
    </div>

    <%
    dim cartid

    cartid = session("cartid")

    if isnull(cartid) then
      response.write "error"
    end if

    dim username

    username = session("username")

    dim t

    dim w

    t = request("t")

    w = request("w")

    if username = "" then
      response.redirect "/login.asp?ru=buy.asp?t=" & t & "&w=" & w
    end if

    dim oRS

    dim sqlstr


    t = formatnumber(t,2)
    %>

  <div id = "container" style = "width:100%;position:relative;top:0px">
      <div id = "middle" style = "float:left; width: 200;">
        Subtotal: $<%=t%><br>
        <div id="sh"></div>
        <script>
        function setShippingCost() {
          var x = document.getElementById("shipping").value;
            x = x.split("$");
            var v;
            v = 0;
            v = x[1];
            v = parseFloat(v,10);
            v = v + <%=t%>;
            v = v.toFixed(2);
            document.getElementById("sh").innerHTML = "Shipping: $" + x[1] + "<br>";
            document.getElementById("sh").innerHTML += "<font color='red' size='5'>Total: $" + v + "</font><br>";
        }
        </script>
      </div>
      <div id = "right" style = "float:left; width: 1250;">
        <form method="POST" action="relayresponse.asp"> 
          <div style="float:left;width:400">
          <h3>Shipping</h3>
          <% Response.Buffer=True %>
<%
sAccessLicenseNumber = "9D5A1C57FAB32998"
sUserID = "dhcho428"
sPassword = "Coppersink21"

dim zip

zip = request.querystring("zip")

if not zip = "" then
  if w = "" then
    w = "1"
  end if
  DrawUPSRates w, zip
end if

Sub DrawUPSRates(sWeight, sDestinationPostalCode)

  sUPSXML = BuildUPSXML(sWeight, sDestinationPostalCode)

  'Now pass the request to UPS
  Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  xmlhttp.Open "POST","https://wwwcie.ups.com/ups.app/xml/Rate",false
  xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  xmlhttp.send sUPSXML

  sResponseXML = xmlhttp.responseText
  Set mydoc=Server.CreateObject("Microsoft.xmlDOM")
  mydoc.loadxml(sResponseXML)
  'Create a select table from the response xml
  response.Write("Shipping Method: <select name='shipping' id='shipping' onchange='setShippingCost()'>")
  'Create A Nodelist of All The RatedShipments
  Set NodeList = mydoc.documentElement.selectNodes("RatedShipment")
  For x = 0 To NodeList.length - 1
    sDisplayString = _
    GetFriendlyUPSName(NodeList.Item(x).selectSingleNode("Service/Code").Text) & _
    " - $" & NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text
    Response.Write("<option>")
    Response.Write (sDisplayString)
    Response.Write("</option>")
  Next
  Response.Write("</select>")
  'response.write(xmlhttp.responseText)
End Sub

Function BuildUPSXML(sWeight, sDestinationPostalCode)
  sShipperPostalCode = "63126"
 
  sXML = sXML & "<?xml version='1.0'?>"
  sXML = sXML & "    <AccessRequest xml:lang='en-US'>"
  sXML = sXML & "        <AccessLicenseNumber>" & sAccessLicenseNumber & _
                "</AccessLicenseNumber>"
  sXML = sXML & "        <UserId>" & sUserID & "</UserId>"
  sXML = sXML & "        <Password>" & sPassword & "</Password>"
  sXML = sXML & "    </AccessRequest>"
  sXML = sXML & "<?xml version='1.0'?>"
  sXML = sXML & "    <RatingServiceSelectionRequest xml:lang='en-US'>"
  sXML = sXML & "        <Request>"
  sXML = sXML & "            <TransactionReference>"
  sXML = sXML & "                <CustomerContext>Rating and " & _
                "Service</CustomerContext>"
  sXML = sXML & "                <XpciVersion>1.0001</XpciVersion>"
  sXML = sXML & "            </TransactionReference>"
  sXML = sXML & "            <RequestAction>Rate</RequestAction>"
  sXML = sXML & "            <RequestOption>shop</RequestOption>"
  sXML = sXML & "        </Request>"
  sXML = sXML & "        <PickupType>"
  sXML = sXML & "            <Code>01</Code>"
  sXML = sXML & "        </PickupType>"
  sXML = sXML & "        <Shipment>"
  sXML = sXML & "            <Shipper>"
  sXML = sXML & "                <Address>"
  sXML = sXML & "                    <PostalCode>" & sShipperPostalCode & _
                "</PostalCode>"
  sXML = sXML & "                </Address>"
  sXML = sXML & "            </Shipper>"
  sXML = sXML & "            <ShipTo>"
  sXML = sXML & "                <Address>"
  sXML = sXML & "                    <PostalCode>" & sDestinationPostalCode & _
                "</PostalCode>"
  sXML = sXML & "                </Address>"
  sXML = sXML & "            </ShipTo>"
  sXML = sXML & "            <Service>"
  sXML = sXML & "                <Code>11</Code>"
  sXML = sXML & "            </Service>"
  sXML = sXML & "            <Package>"
  sXML = sXML & "                <PackagingType>"
  sXML = sXML & "                    <Code>02</Code>"
  sXML = sXML & "                    <Description>Package</Description>"
  sXML = sXML & "                </PackagingType>"
  sXML = sXML & "                <Description>Rate Shopping</Description>"
  sXML = sXML & "                <PackageWeight>"
  sXML = sXML & "                    <Weight>" & sWeight & "</Weight>"
  sXML = sXML & "                </PackageWeight>"
  sXML = sXML & "            </Package>"
  sXML = sXML & "            <ShipmentServiceOptions/>"
  sXML = sXML & "        </Shipment>"
  sXML = sXML & "</RatingServiceSelectionRequest>"
 
  BuildUPSXML = Replace(sXML, vbTab, "")
End Function

Function GetFriendlyUPSName(vCode)
  Select Case vCode
    Case "01"
      GetFriendlyUPSName = "UPS Next Day Air"
    Case "02"
      GetFriendlyUPSName = "UPS 2nd Day Air"
    Case "03"
      GetFriendlyUPSName = "UPS Ground"
    Case "07"
      GetFriendlyUPSName = "UPS Worldwide Express"
    Case "08"
      GetFriendlyUPSName = "UPS Worldwide Expedited"
    Case "11"
      GetFriendlyUPSName = "UPS Standard"
    Case "12"
      GetFriendlyUPSName = "UPS 3 Day Select"
    Case "13"
      GetFriendlyUPSName = "UPS Next Day Air Saver"
    Case "14"
      GetFriendlyUPSName = "UPS Next Day Air Early A.M."
    Case "54"
      GetFriendlyUPSName = "UPS Worldwide Express Plus"
    Case "59"
      GetFriendlyUPSName = "UPS 2nd Day Air A.M."
    Case "65"
      GetFriendlyUPSName = "UPS Saver"
  End Select
End Function

%>
          <script type="text/javascript">
            setShippingCost();
          </script>
          <br>
          <script>
            var qs = window.location.search;
            if(qs.indexOf('index=1') > -1) {
              document.write("<font color=red>Your address could not be verified.</font><br>");
            }
          </script>
          Your Address 1: <input type=text id="x_custom_3" onclick="document.getElementById('steps').innerHTML='<img src=/graphics/images/step1.png><br><br>'" name="x_custom_3"/><br><br>
          Your Address 2 (Suite, Apt, etc.): <input type=text id="x_custom_8" name="x_custom_8"/><br><br>
          Your City: <input type=text id="x_custom_4" name="x_custom_4"/><br><br>
          Your State: <input type=text id="x_custom_5" name="x_custom_5"/><br><br>
          Your Zip Code: <input type=text id="x_custom_6" name="x_custom_6" value="<%=zip%>"/><br><span style="background-color:green;color:white;"><a onclick="window.location.href='buy.asp?t=<%=t%>&w=<%=w%>&zip='+document.getElementById('x_custom_6').value">Set for shipping method.</a></span><br><br>
          Your Country: <input type=text id="x_custom_7" name="x_custom_7"/><br><br><br>
          </div>
          <div style="float:left;width:400">
          <h3>Billing</h3>
          Your Address: <input type=text id="address" onclick="document.getElementById('steps').innerHTML='<img src=/graphics/images/step2.png><br><br>'" name="address"/><br><br>
          Your City: <input type=text id="city" name="city"/><br><br>
          Your State: <input type=text id="state" name="state"/><br><br>
          Your Zip Code: <input type=text id="zipcode" name="zipcode"/><br><br>
          Your Country: <input type=text id="country" name="country"/><br><br><br>
          <input type=hidden id="x_login" name="x_version" value='3.1'/>
          <input type=hidden id="x_login" name="x_delim" value='False'/>
          <input type=hidden id="x_login" name="x_login" value='6r46jMVd4X'/>
          <input type=hidden id="x_tran_key" name="x_tran_key" value='84XDv9Xq2E7ay6s7'/>
          </div>
          <div style="float:left;width:400">
          <h3>Credit Card</h3>
          Your Card Type: 
          <select name="cc">
            <option value="visa">VISA</option>
            <option value="mc">MASTER CARD</option>
            <option value="dsc">DISCOVER</option>
            <option value="amx">AMEX</option>
          </select><br><br>
          Your Credit Card Number: <input type=text id="x_card_num" onclick="document.getElementById('steps').innerHTML='<img src=/graphics/images/step3.png><br><br>'" name="x_card_num"/><br><br>
          Your Exp Date: <input type=text id="x_exp_date" name="x_exp_date"/><br><br>
          Your Security Code: <input type=text id="x_card_code" name="x_card_code"/><br><br>
          <%
          if t <> "" and w <> "" and zip <> "" then
          %>
          <input type="submit" value="Proceed to Pay" style="width: 150px; height:50px" />
          <%
          end if
          %>
          <input type=hidden id="x_amount" name="x_amount" value='<%=t%>' />
          <input type=hidden id="x_relay_url" name="x_relay_url" value='https://pricebay.gq/relayresponse.asp'/>
          <input type=hidden id="x_relay_response" name="x_relay_response" value='Y'/>
<input type='hidden' name="x_show_form" value="payment_form">
<input type='hidden' name="x_test_request" value="false" />
<input type='hidden' name="x_method" value="cc">
          <input type=hidden id="x_type" name="x_type" value='AUTH_CAPTURE'/>
          <input type=hidden id="x_currency_code" name="x_currency_code" value='USD'/>
          <input type=hidden name="x_custom_1" value='<%=session("username")%>' />
          <input type=hidden name="x_custom_2" value='<%=cartid%>' />
          </div>
        </form>
      </div>
  </div>
  <%
  %>

  </div>

<!--#include file="inc/footer.inc"-->