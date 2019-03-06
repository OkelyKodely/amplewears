<!--#include file="inc/header.inc"-->

  <script language="javascript">

    function interpretUsername(username) {
  
      return username.value.length > 5;
  
    }

    function interpretEmail(email) {
  
      if (/^(\S+)@(\S+).(\S+)/.exec(email)) {
  
        return true;
  
      } else {
  
        return false;
  
      }
  
    }

    function validateThenSubmit() {
  
      var frm = document.getElementById("frm");

      var username = document.getElementById("username");
  
      var usernameIsValid = interpretUsername(username);

      var email = document.getElementById("email").value;
  
      var emailIsValid = interpretEmail(email);

      if(usernameIsValid &&
         emailIsValid) {
  
        frm.submit();
  
      } else {
  
        alert("fix your errors!!~");
  
      }
  
    }
  
  </script>

  <%
  dim firstname
  dim lastname
  dim username
  dim password
  dim pusername
  dim address
  dim city
  dim stateorprovince
  dim country
  dim phone
  dim fax
  dim email
  dim subscribe

  firstname = request.form("firstname")
  lastname = request.form("lastname")
  username = request.form("username")
  password = request.form("password")
  pusername = request.form("pusername")
  address = request.form("address")
  city = request.form("city")
  stateorprovince = request.form("stateorprovince")
  country = request.form("country")
  phone = request.form("phone")
  fax = request.form("fax")
  email = request.form("email")
  subscribe = request.form("subscribe")

'dim sqlstr1
'sqlstr1 = "CREATE TRIGGER shopp_bi BEFORE INSERT ON shoppers2 FOR EACH ROW BEGIN SELECT shopp_seq.nextval INTO :new.userid FROM dual; END;"
'Conn.Execute(sqlstr1)

  if firstname <> "" then
    dim sqlstr

    dim oRs

    sqlstr = "SELECT * FROM shoppers2 WHERE email = '"&email&"'"

    set oRs = Conn.Execute(sqlstr)

    if oRs.eof then

      if subscribe = "checked" then

        subscribe = 1

      else

        subscribe = "0"

      end if

      sqlstr = "INSERT INTO shoppers2 (subscribe, firstname, lastname, username, password, address, city, stateorprovince, country, phone, fax, email, paypal_user) VALUES ("&subscribe&",'"&firstname&"','"&lastname&"','"&username&"','"&password&"','"&address&"','"&city&"','"&stateorprovince&"','"&country&"','"&phone&"','"&fax&"','"&email&"','"&pusername&"')"
'response.write sqlstr
      Conn.Execute(sqlstr)

      response.write "Registered!"

    else
    %>
    <script>
      alert('Account for email <%=email%> already exists.');
    </script>
    <%
    end if

  end if
  %>

  <div style="height:1000px;float:left;position:relative;top:100px">
    <div style="height:80px">
      <div>
      <span style="font-size:40px;">Sign Up - Create Account</span><br>
      <span><br></span>
      <span><br></span>
    </div>
    </div>
    <form method="post" id="frm" action="" onsubmit="return validateThenSubmit()">
      <div style="float:left;width:140px">First Name:</div>
      <div style="loat:left"><input type="text" name="firstname"></div>
      <div style="float:left;width:140px">Last Name:</div>
      <div style="loat:left"><input type="text" name="lastname"></div>
      <div style="float:left;width:140px">Username:</div>
      <div style="loat:left"><input type="text" id="username" name="username"></div>
      <div style="float:left;width:140px">Password:</div>
      <div style="loat:left"><input type="password" name="password"></div>
      <div style="float:left;width:140px">Paypal Username:</div>
      <div style="loat:left"><input type="text" id="pusername" name="pusername"></div>
      <div style="float:left;width:140px">Address:</div>
      <div style="loat:left"><input type="text" name="address"></div>
      <div style="float:left;width:140px">City:</div>
      <div style="loat:left"><input type="text" name="city"></div>
      <div style="float:left;width:140px"><nobr>State or Province:</nobr></div>
      <div style="loat:left"><input type="text" name="stateorprovince"></div>
      <div style="float:left;width:140px">Country:</div>
      <div style="loat:left"><input type="text" name="country"></div>
      <div style="float:left;width:140px">Phone:</div>
      <div style="loat:left"><input type="text" name="phone"></div>
      <div style="float:left;width:140px">Fax:</div>
      <div style="loat:left"><input type="text" name="fax"></div>
      <div style="float:left;width:140px">Email:</div>
      <div style="loat:left"><input type="text" id="email" name="email"></div>
      <div style="float:left;width:140px">Subscribe:</div>
      <div style="loat:left"><input type="checkbox" id="subscribe" name="subscribe"> yes?</div>
      <p style="font-size:11px">You are given the option to subscribe to our marketing material.  <br>By checking "yes" on Subscribe you agree to receive information from us <br>regarding our products and services from time to time.</p>
      <br><br>
      <input type="submit" value="Sign Up" style="width:350px; height:50px; position: relative; left: 0px">
    </form>
  </div>
<!--#include file="inc/footer.inc"-->