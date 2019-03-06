<!--#include file="inc/header.inc"-->
<%
  dim username
  dim password

  if session("username") = "" then

    username = request.form("username")
    password = request.form("password")

  else

    username = session("username")
    password = session("password")

  end if

  %>
  <div class="login" style="float:left; width:100%; height:700px;text-align:center">
  <h1 style="position:relative;left:0px;top:100px">Returning Customer - Log On</h1>
  <%
  if username = "" then
  %>
      <div style="width:400px;height:600px;background-image:url('/graphics/images/loginbg.png');background-repeat:no-repeat;background-attached:fixed;background-position:center;">
        <div>

    <div style="text-align:center;position:relative;left:350px;top:100px;z-index:2">
    <div style="height:70px;position:relative;top:0px">
    </div>
          <form method="post" action="">
            <div style="float:left;width:90px">Username:</div>
            <div style="loat:left"><input type="text" name="username"></div>
            <span><br></span>
            <div style="float:left;width:90px">Password:</div>
            <div style="loat:left"><input type="password" name="password"></div>
            <input type="submit" value="Login" style="top:30px;width:140px;height:40px;position:relative;left:0px">
          </form>
        </div>
      </div>
    </div>

      <div style="float:left;width:400px;height:600px;ackground-image:url('/graphics/images/loginbg.png');background-repeat:no-repeat;background-attached:fixed;background-position:center;position:relative;">
        <div>

    <div style="position:relative;left:600px;top:-400px;z-index:2">
    <div style="height:70px;position:relative;top:0px">
    </div>
    <p style="color:black">Don't have an account with us?</p>
    <a href="signup.asp"><h1 style="color:black">Create Account</h1></a>
        </div>
      </div>
    </div>

<%
else

    dim sqlstr

    sqlstr = "SELECT * FROM shoppers2 WHERE username='"&username&"' AND password='"&password&"'"

    set oRS = Conn.Execute(sqlstr)

if not oRS.eof then
  session("username") = oRS("username")
%><br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=oRS("username")%> logged on.  Thanks.
<%
    dim ru
    ru = request("ru")
    if ru <> "" then
      response.redirect ru
    end if
  else
    if session("username") = "" then
      response.write "not logged on"
    end if
  end if
end if
%>
  </div>

<!--#include file="inc/footer.inc"-->