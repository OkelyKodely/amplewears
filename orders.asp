<!--#include file="inc/header.inc"-->
  <div>
  
    <div style="position:relative;top:140px">
      <span style="font-size:26;color:black">Your Order History With AmpleWears</span>
    </div>

    <p>
      &nbsp;
    </p>

    <p>
      &nbsp;
    </p>

    <%
    dim oRS

    dim sqlstr

    dim username

    username = session("username")

    dim a, b

    a = request("a")

    b = request("b")

    if a = "" then
    a = 1
    b = 6
    end if

    dim ssql

    ssql = "SELECT * FROM   (    SELECT *    FROM   ""order2"" where username = '"&username&"' order by orderid desc) WHERE rownum between "&a&" AND "&b&";"

    set oRS = Conn.Execute(ssql)

    %>

          <div style = "float:left; width: 50px; height: 30px;position:relative;top:100px">
      <a href='orders.asp?a=<%=a-6%>&b=<%=b-6%>' style="color:black">less...</a>
    </div>

    <%

    dim count
    count = 0
    while not oRS.eof
      %>
      <div style = "width:100%;float:left;color:black;position:relative;top:100px">
        <div stylef = "width:800px;float:left;position:relative;left:-100px">
          <div style="float:left;width:180px"><a style="color:black" href='orderitems.asp?orderid=<%=oRS("orderid")%>'%><%=oRS("orderid")%></div>
          <div style="float:left;width:120px;color:black">$<%=oRS("price")%> USD</div>
          <div style="color:black">
            <nobr>
              <%=oRS("inputdate")%>
            </nobr>
          </div>
        </div><br>
      </div>
      <%
      count = count + 1
      oRS.movenext
    wend

Conn.close()
%>

      <div style = "float:left; width: 100%; height: 30px;position:relative;top:100px">
        <a href='orders.asp?a=<%=a+6%>&b=<%=b+6%>' style="color:black">more...</a>
      </div>
    </div>

        <p>
      &nbsp;
    </p>

    <p>
      &nbsp;
    </p>

    <p>
      &nbsp;
    </p>

    <p>
      &nbsp;
    </p>

<!--#include file="inc/footer.inc"-->