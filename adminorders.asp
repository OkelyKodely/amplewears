<style>
a {
  color:black;
}
</style>
<%
  server.scripttimeout = 1000

  dim username
  dim password

  if session("ausername") = "" then

    response.redirect "adminlogin.asp"

  else

    username = session("ausername")
    
  end if

Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

  %>
  <div style="float:left; width:100%">
  <%
  if username = "" then

  end if

    %>
    <a href="adminitems.asp">Admin Items</a><Br /><Br />
    <table>
      <tr>
        <td style="border-bottom:1px solid gray" width="200">order id</td>
        <td style="border-bottom:1px solid gray" width="200">amount</td>
        <td style="border-bottom:1px solid gray" width="200">shipping method</td>
        <td style="border-bottom:1px solid gray" width="200">shipping cost</td>
        <td style="border-bottom:1px solid gray" width="200">status</td>
        <td style="border-bottom:1px solid gray" width="200">input date</td>
    <%

  if username = "admin_cho4321" then
  session("ausername") = username
    sqlstr = "SELECT * FROM ""order2"" order by orderid desc"

    set rs = Conn.Execute(sqlstr)
    while not rs.eof
  %>
    <tr>
        <td><a href='adminorderitems.asp?orderid=<%=rs("orderid")%>'><%=rs("orderid")%></a></td>
        <td><%=rs("price")%></td>
        <td><%=rs("shippingmethod")%></td>
        <td><%=rs("shippingcost")%></td>
        <td><%=rs("status")%></td>
        <td><%=rs("inputdate")%></td>
    </tr>
  <%
      rs.movenext
    wend
  else
    if session("ausername") = "" then
      response.write "not signed in"
    end if
  end if
%>
    </table>
  </div>

<%
Conn.close()
%>