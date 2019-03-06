
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

dim itemid

itemid = request.querystring("itemid")

  %>
  <div style="float:left; width:100%">
  <%
  if username = "" then

  end if

    %>
    <a href="adminitems.asp">Admin Items</a>
    <%

  if username = "admin_cho4321" then
  session("ausername") = username
    sqlstr = "SELECT weight, bestseller,image,itemid,name,price,discount,category,newarrivals,dsc FROM products2 WHERE itemid = " & itemid
'response.write sqlstr
    set rs = Conn.Execute(sqlstr)
    if not rs.eof then
  %>
  <form method="post" action="upload2.asp" enctype="multipart/form-data">
  <input type="hidden" name="itemid" value='<%=rs("itemid")%>'>
  <input type="hidden" name="up" value="yes">
  <%
  if rs("newarrivals") = "1" then
  %>
  <input type="text" name="newarrival" value="on"> New Arrival?<br>
  <%
  else
  %>
  <input type="text" name="newarrival" value="off"> New Arrival?<br>
  <%
  end if
  %>
  <%
  if rs("bestseller") = "1" then
  %>
  <input type="text" name="bestseller" value="on"> Best Seller?<br>
  <%
  else
  %>
  <input type="text" name="bestseller" value="off"> Best Seller?<br>
  <%
  end if
  %>
  Name: <input type="text" name="name" value='<%=rs("name")%>'><br>
  Category: <input type="text" name="category" value='<%=rs("category")%>'><br>
  Recommendation: <input type="text" name="theitemid"><br>
  Price: <input type="text" name="price" value='<%=rs("price")%>'><br>
  Weight: <input type="text" name="weight" value='<%=rs("weight")%>'><br>
  Image: <input type="file" name="file"><br>
  Women: <input type="text" name="women"><br>
  Men: <input type="text" name="men"><br>
  Discount: <input type="text" name="discount"><br>
  <img src='/productitems/<%=rs("image")%>'><br>
  Description: <textarea name="dsc" style="width:300px;height:100px"><%=rs("dsc")%></textarea><br>
  <input type="submit" name="submit">
</form>
  <%
    end if
  else
    if session("ausername") = "" then
      response.write "not signed in"
    end if
  end if
%>
  </div>

<h1>&nbsp;</h1>
<%
Conn.close()
%>