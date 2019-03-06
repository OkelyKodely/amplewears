
  <%
  server.scripttimeout = 1000
  
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
    %>
    <a href="adminitems.asp">Admin Items</a> <a href="adminitemsinsert.asp">Items Insert</a>
    <%
  end if

  if username = "admin_cho4321" then
  session("ausername") = username
  %>
  <form method="post" action="upload1.asp" enctype="multipart/form-data">
  <input type="hidden" name="up" value="yes">
  <input type="text" name="newarrival"> New Arrival?<br>
  <input type="text" name="bestseller"> Best Seller?<br>
  Name: <input type="text" name="name"><br>
  Category: <input type="text" name="category"><br>
  Recommendation: <input type="text" name="theitemid"><br>
  Price: <input type="text" name="price"><br>
  Discount: <input type="text" name="discount"><br>
  Weight (Lbs): <input type="text" name="weight"><br>
  Women: <input type="text" name="women"><br>
  Men: <input type="text" name="men"><br>
  Featured: <input type="text" name="featured"><br>
  Description: <textarea name="dsc" style="width:300px;height:100px"></textarea><br>
  Image: <input type="file" name="file">
    <br/><br/>
    <input type="submit" name="submit">
</form>
  <%
  else
    if session("ausername") = "" then
      response.write "not signed in"
    end if
  end if
%>
  </div>

<h1>&nbsp;</h1>