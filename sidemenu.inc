<div style="width:230px;text-align:center;position:relative;left:10px;top:10px">
  <div style="width:100%;border:1px solid gray;height:40px">
  <a href="index.asp" style="color:white;font-size:20px;text-decoration:underlin">
    Home
  </a>
  </div>
  <div style="width:100%;border:1px solid gray;height:40px">
  <a href="about.asp" style="color:white;font-size:20px;text-decoration:underlin">
    About
  </a>
  </div>
  <div style="width:100%;border:1px solid gray;height:40px">
  <a href="cart.asp" style="color:white;font-size:20px;text-decoration:underlin">
    Cart
  </a>
  </div>
  <div style="width:100%;border:1px solid gray;height:40px">
      <%
      if session("username") = "" then
      %>
  <a href="signup.asp" style="color:white;font-size:20px;text-decoration:underlin">
    Register
  </a>
  <%
else
  %>
  <a href="orders.asp" style="color:white;font-size:20px;text-decoration:underlin">
    Orders
  <%
end if
  %>
  </div>
  <div style="width:100%;border:0.05em solid gray;height:40px">
  <a href="contactus.asp" style="color:white;font-size:20px;text-decoration:underlin">
    Contact Us
  </a>
  </div>
      <%
      if session("username") = "" then
      %>
      <form id="loginFrm" method="post" action="login.asp">
        <div style="color:white;">
      <h2>Sign in</h2>
        Username: <br><input type="text" name="username"><br>
        Password: <br><input type="password" name="password">
        <input type="hidden" name="ru" value="shop.asp"><br>
        <button class="favorite styled3" style="width:200px;color:black;" onclick="document.getElementById('loginFrm').submit()">Sign in</button>
        <br><br>
        <h4>New to AmpleWears? <br><a style="color:gray" href="signup.asp">Create Account</a></h4>
        <%
        else
        %>
        <br><br>
        <h3 style="color:gray">Hi, <%=session("username")%>.<br> Welcome.</h3>
        <br>
        <ul style="width:100px;position:relative;left:35px">
        <li><a href="orders.asp" style="color:gray">Your Orders</a><br><br></li>
        <li><a href="cart.asp" style="color:gray">Your Cart</a><br><br></li>
        <li><a href="logout.asp" style="color:gray">Sign out</a></li>
      </ul>
        <%
        end if
        %>
        </div>
</div>