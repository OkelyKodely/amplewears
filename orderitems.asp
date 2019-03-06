<!--#include file="inc/header.inc"-->
    <div>
      <span style="font-size:26;color:black">Your <span onclick="window.location.href='orders.asp'"><u>Order History</u></span> With AmpleWears</span>
    </div>

    <p>
      &nbsp;
    </p>

    <p>
      &nbsp;
    </p>

    <div style="width:650px">
    <%
    dim oRS

    dim sqlstr

    dim orderid

    orderid = request.querystring("orderid")

    dim a, b

    a = request("a")

    b = request("b")

    if a = "" then
    a = 1
    b = 6
    end if

    dim ssql

    ssql = "SELECT * FROM  (    SELECT a.itemid, p.price, p.category, p.image, p.discount "
    ssql = ssql & "    FROM orderitems2 a INNER JOIN products2 p ON a.itemid = p.itemid WHERE a.orderid="&orderid&") WHERE rownum between "&a&" AND "&b&";"

    set oRS = Conn.Execute(ssql)

    response.write "<h1>Order ID: " & orderid & "</h1>"
    %>

    <div style = "width:100%;float:left">

    <div style="float:left;width:180px;color:black">item id</div>
    <div style="float:left;width:120px;color:black">price</div>
    <div style="float:left;width:120px;color:black">category</div>
    <div style="float:left;width:120px;color:black">image</div>
    
    </div>
    
    <div style = "width:800px;float:left">
      <span><hr style="color:#f0f0f0"></span>
    </div>
          <div style = "float:left; width: 50px; height: 30px">
      <a href='orderitems.asp?orderid=<%=orderid%>&a=<%=a-6%>&b=<%=b-6%>' style="color:black">less...</a>
    </div>
    <%

    while not oRS.eof
      %>
      <div style = "width:100%;float:left">

      <div style="float:left;width:180px"><a href='item.asp?itemid=<%=oRS("itemid")%>' style="color:black"><%=oRS("itemid")%></a></div>
      <div style="float:left;width:120px;color:black">$<%=formatnumber((100-cdbl(oRS("discount")))*cdbl(oRS("price"))/100,2)%> USD</div>
      <div style="float:left;width:120px;color:black"><%=oRS("category")%></div>
      <div style="float:left;width:120px;color:black"><img src='/productitems/<%=oRS("image")%>'></div>

      </div><br>
      <%
      oRS.movenext
    wend
    %>

    </div>
      <div style = "float:left; width: 50px; height: 30px">
      <a href='orderitems.asp?orderid=<%=orderid%>&a=<%=a+6%>&b=<%=b+6%>' style="color:black">more...</a>
      </div>

  </div>

    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
<%
Conn.close()
%>

<!--#include file="inc/footer.inc"-->