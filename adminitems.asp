
  <%
  dim username
  dim password

  if session("ausername") = "" then

    response.redirect "adminlogin.asp"

  else

    username = session("ausername")
    
  end if

  %>
  <div style="float:left; width:100%">
  <%
  if username = "" then
  else

    dim sqlstr

Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

sqlstr = "select count(*) as cnt from products2"

dim rs

set rs = Conn.execute(sqlstr)

count = rs("cnt")

sqlstr = "select * from products2 order by itemid desc"

set rs = Conn.execute(sqlstr)
    %>
    <a href="adminitemsinsert.asp">Items Insert</a><br>
    Total Items: <%=count%><br>
    <table width="100%">
      <tr>
        <td>Name</td>
        <td>Sex</td>
        <td>Category</td>
        <td>Price</td>
      </tr>
      <tr>
        <td colspan="4">
          <hr>
        </td>
      </tr>
      </tr>
    <%

  if username = "admin_cho4321" then
  session("ausername") = username
  while not rs.eof
%>
  <tr>
    <td><a href='adminitemsedit.asp?itemid=<%=rs("itemid")%>'><%=rs("name")%></a></td>
    <td><%if rs("men") = 1 then response.write "men"%><%if rs("women") = 1 then response.write "women"%></td>
    <td><%=rs("category")%></td>
    <td><%=rs("price")%></td>
  </tr>
<%
  rs.movenext
  wend
  %>
  </table>
  <%
  else
    if session("ausername") = "" then
      response.write "not signed in"
    end if
  end if
  end if
%>
  </div>

<h1>&nbsp;</h1>
<%
Conn.close()
%>