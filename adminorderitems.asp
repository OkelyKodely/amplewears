<style>
a {
  color:black;
}
</style>
<%
dim username
dim password

if session("ausername") = "" then

  response.redirect "adminlogin.asp"

else

  username = session("ausername")
    
end if

dim orderid, task, ssql, transid, cc4
orderid = request("orderid")
task = request("task")

Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

    ssql = "SELECT b.email as fool, a.transid, a.cc4, a.username, a.status, a.price, a.shippingcost, a.shippingmethod, a.inputdate FROM ""order2"" a inner join shoppers2 b on a.username = b.username where a.orderid = " & orderid
    set oRS = Conn.Execute(ssql)

    dim total, shipmentcost, shipmentmethod, orderdate, mailtothisfool
    if not oRS.eof then
      total = oRS("price")
      shipmentcost = oRS("shippingcost")
      shipmentmethod = oRS("shippingmethod")
      orderdate = oRS("inputdate")
      mailtothisfool = oRS("fool")
      transid = oRS("transid")
      cc4 = oRS("cc4")
    end if

    if task <> "" then

      dim continu

      if task = "refunded" then

        Dim post_url
        post_url = "https://test.authorize.net/gateway/transact.dll"

        Dim x_login, x_tran_key
        x_login = "6r46jMVd4X"
        x_tran_key = "84XDv9Xq2E7ay6s7"
                  
        Dim post_values
        Set post_values = CreateObject("Scripting.Dictionary")
        post_values.CompareMode = vbTextCompare

        post_values.Add "x_login", x_login
        post_values.Add "x_tran_key", x_tran_key

        post_values.Add "x_delim_data", "TRUE"
        post_values.Add "x_delim_char", ","
        post_values.Add "x_relay_response", "FALSE"

        post_values.Add "x_type", "CREDIT"
        post_values.Add "x_amount", formatnumber(cdbl(total)*(1.00-0.15), 2)
        post_values.Add "x_card_num", cc4
        post_values.Add "x_trans_id", transid

       '' response.write formatnumber(cdbl(total)*(1.00-0.15), 2) & "<BR>"
       '' response.write cc4 & "<BR>"
       '' response.write transid & "<BR>"

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
        continu = 0

 ''''response.write post_response

        dim cnt
        cnt = 0
        dim value
        For Each value in response_array
          if cnt = 0 then
            if value = "1" then
              continu = 1
            end if
          end if
          cnt = cnt + 1
        Next

      end if

      if task <> "refunded" or continu = 1 then

        ssql = "update ""order2"" set status = '"&task&"' where orderid = " & orderid
        Conn.Execute(ssql)

        Dim objMail 
        Set objMail = Server.CreateObject("CDO.Message") 

        Dim smtpServer, yourEmail, yourPassword
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

        if task = "shipping..." then
          objMail.Subject="Your AmpleWears Order has shipped"
          objMail.htmlBody = "Your order in the amount of $" & total & " USD has shipped.  Thank you for doing business with us.<br><br>AmpleWears team"
        elseif task = "done..." then
          objMail.Subject="Your AmpleWears Order has been delivered"
          objMail.htmlBody = "Your order in the amount of $" & total & " USD has been delivered.  Thank you for doing business with us.<br><br>AmpleWears team"
        elseif task = "refunded" then
          objMail.Subject="Your AmpleWears Order has been refunded"
          objMail.htmlBody = "Your order in the amount of $" & total & " USD has been 85% partially refunded with the 15% going to doing stocking fee (for a total partial refund of $"&formatnumber(cdbl(total)*(1.00-0.15), 2)&").  We look forward to doing business with you again soon.<br><br>AmpleWears team"
        end if
        objMail.From = "sale.pbay@gmail.com"
        objMail.To = mailtothisfool
        objMail.Send

      end if

    end if

    if not oRS.eof then
      if task = "" then
        task = oRS("status")
      end if
    end if

    ssql = "SELECT * FROM  (    SELECT a.itemid, p.price, p.category, p.image, p.discount "
    ssql = ssql & "    FROM orderitems2 a INNER JOIN products2 p ON a.itemid = p.itemid WHERE a.orderid="&orderid&");"

    set oRS = Conn.Execute(ssql)

    response.write "<h1>Order ID: " & orderid & "</h1>"
    %>
    <div>
      <a href='adminorders.asp'>Back to Orders</a>
    </div>

    <Br />

    <div>
      <a href='admincreatelabel.asp?orderid=<%=orderid%>'>Create UPS Shipment Label</a>
    </div>

    <div>
      Total: $<%=total%><br>
      Shipping Cost: $<%=shipmentcost%><br>
      Shipping Method: <%=shipmentmethod%><br>
      <%=orderdate%><br>
    </div>

    <div>
      <form method="post" action="">
        <select name="task" onchange="this.form.submit()">
          <option <%if task = "processing..." then response.write "selected"%>>processing...</option>
          <option <%if task = "shipping..." then response.write "selected"%>>shipping...</option>
          <option <%if task = "done..." then response.write "selected"%>>done...</option>
          <option <%if task = "refunded" then response.write "selected"%>>refunded</option>
        </select>
        <input type="hidden" name="orderid" value='<%=orderid%>'>
      </form>
    </div>

    <Br />

    <div style = "width:100%;float:left">

    <div style="float:left;width:180px">item id</div>
    <div style="float:left;width:120px">price</div>
    <div style="float:left;width:120px">category</div>
    <div style="float:left;width:120px">image</div>
    
    </div>
    
    <div style = "width:800px;float:left">
      <span><hr style="color:#f0f0f0"></span>
    </div>
    <%

    while not oRS.eof
      %>
      <div style = "width:100%;float:left">

      <div style="float:left;width:180px"><a href='item.asp?itemid=<%=oRS("itemid")%>'><%=oRS("itemid")%></a></div>
      <div style="float:left;width:120px">$<%=formatnumber((100-cdbl(oRS("discount")))*cdbl(oRS("price"))/100,2)%> USD</div>
      <div style="float:left;width:120px"><%=oRS("category")%></div>
      <div style="float:left;width:120px"><img src='/productitems/<%=oRS("image")%>'></div>

      </div><br>
      <%
      oRS.movenext
    wend
    %>

<%
Conn.close()
%>