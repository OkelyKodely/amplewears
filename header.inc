<%
dim Conn
Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"
dim cat
cat = request("cat")
%>
<html>
  <head>
    <title>The Official Website of Ample Wears</title>
    <link rel="stylesheet" href="bootstrap.css" rel="stylesheet">
  </head>
  <body leftmargin="0" topmargin="0">
  	<div style="z-index:7;float:left;padding:20px 20px 20px 20px">
  	  <a href="/index.asp"><img src="/graphics/images/logo.png"></a>
  	</div>

<style>
body {font-family: Arial;color:black;}

/* Style the tab */
.tab {
  overflow: hidden;
  border: 0px solid #ccc;
  background-color: #fff;
}

a {
	color:white;
	text-decoration: none;
	font-size:12px;
}

/* Style the buttons inside the tab */
.tab button {
  background-color: inherit;
  float: left;
  border: none;
  outline: none;
  cursor: pointer;
  padding: 14px 16px;
  transition: 0.3s;
  font-size: 17px;
}

/* Change background color of buttons on hover */
.tab button:hover {
  background-color: #ddd;
}

/* Create an active/current tablink class */
.tab button.active {
  background-color: #ccc;
}

/* Style the tab content */
.tabcontent {
  display: none;
  padding: 6px 12px;
  border: 0px solid #ccc;
  border-top: none;
}
</style>
<style>
.dropbtn {
  color: white;
  border: none;
}

.dropdown {
  position: relative;
  display: inline-block;
}

.dropdown-content {
  display: none;
  position: absolute;
  background-color: #f1f1f1;
  min-width: 160px;
  box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
  z-index: 1;
}

.dropdown-content a {
  color: black;
  padding: 12px 16px;
  text-decoration: none;
  display: block;
}

.dropdown-content a:hover {color:black;}

.dropdown:hover .dropdown-content {display: block;}

.dropdown:hover .dropbtn {}

a:hover {
  color:white;
  text-decoration:none;
}
      #myBtn {
        display: none; /* Hidden by default */
        position: fixed; /* Fixed/sticky position */
        bottom: 20px; /* Place the button at the bottom of the page */
        right: 30px; /* Place the button 30px from the right */
        z-index: 99; /* Make sure it does not overlap */
        border: 1px solid black; /* Remove borders */
        outline: none; /* Remove outline */
        background-color: #fff; /* Set a background color */
        color: #000; /* Text color */
        cursor: pointer; /* Add a mouse pointer on hover */
        padding: 15px; /* Some padding */
        border-radius: 15px; /* Rounded corners */
        font-size: 18px; /* Increase font size */
      }

      #myBtn:hover {
        color:black;
        background-color: #fff;
      }
</style>

    <script>
// When the user scrolls down 20px from the top of the document, show the button
window.onscroll = function() {scrollFunction()};

function scrollFunction() {
  if (document.body.scrollTop > 20 || document.documentElement.scrollTop > 20) {
    document.getElementById("myBtn").style.display = "block";
  } else {
    document.getElementById("myBtn").style.display = "none";
  }
}

// When the user clicks on the button, scroll to the top of the document
function topFunction() {
  document.body.scrollTop = 0; // For Safari
  document.documentElement.scrollTop = 0; // For Chrome, Firefox, IE and Opera
} 
    </script>

<div class="tab" style="padding:20px 20px 0px 0px">
  <button class="tablinks" style="border-right:1px solid white;background-color:black;color:white" onclick="openCity(event, 'London')" id="defaultOpen">Women</button>
  <button id="men" class="tablinks" style="background-color:#eee;color:black" onclick="openCity(event, 'Paris')">Men</button>
  <div style="text-align:center;float:left;height:20px;width:400px">
  <form method="post" action="searchresults.asp">
    <input type="text" name="search" style="width:300px;height:34px">
    <input type="submit" value="Search">
  </form>
</div>
  <a href="cart.asp"><img width="70" height="30" src="/graphics/images/cart.png"></a>
  <%
  if session("username") = "" then
  %>
  <a href="login.asp" style="color:black">Sign in</a><img src="/graphics/images/downs2.png"> 
  <a href="signup.asp" style="color:black">Register</a><img src="/graphics/images/downs2.png"> 
  <%
  else
  %>
  Hi <%=session("username")%>. Welcome.  <a href="logout.asp" style="color:gray">Sign out</a>
  <%
  end if 
  %>
</div>

<div id="London" class="tabcontent" style="text-align:center;position:absolute;left:0px;background-color:black;color:white;padding:10px 10px 10px 10px">
  <div class="dropdown">
  <a class="dropbtn" href="newarrivals.asp?s=women">NEW ARRIVALS</a>
  <div class="dropdown-content">
    <a href="#">MARIGOLD COLLECTION</a>
    <a href="#">NEW DRESSES</a>
    <a href="#">NEW LINGERIE</a>
    <a href="#">NEW JEANS</a>
    <a href="#">NEW TOPS</a>
    <a href="#">NEW GRAPHIC TOPS</a>
    <a href="#">NEW JACKETS</a>
    <a href="#">NEW SWEATERS</a>
    <a href="#">NEW BOTTOMS</a>
    <a href="#">NEW MATCHING SETS</a>
    <a href="#">NEW SHOES</a>
    <a href="#">NEW ACCESSORIES</a>
    <a href="#">NEW PLUS</a>
    <a href="#">NEW ACTIVEWEAR</a>
  </div>
</div>
   <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=dresses">DRESSES</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">LITTLE BLACK DRESS</a>
    <a href="#">LITTLE WHITE DRESS</a>
    <a href="#">LITTLE RED DRESS</a>
    <a href="#">LUXE</a>
    <a href="#">SPECIAL OCCASION</a>
    <a href="#">MINI DRESSES</a>
    <a href="#">MIDI DRESSES</a>
    <a href="#">MAXI DRESSES</a>
    <a href="#">EVERYDAY DRESSES</a>
    <a href="#">$15 & UNDER DRESSES</a>
    <a href="#">SALE DRESSES</a>
  </div>
  </div>
   <img src="/graphics/images/downs.png">
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=jeans">JEANS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">BEST SELLING JEANS</a>
    <a href="#">HIGH WAIST</a>
    <a href="#">MID RISE</a>
    <a href="#">SKINNY</a>
    <a href="#">BOOTCUT/FLARE</a>
    <a href="#">BOYFRIEND</a>
    <a href="#">DISTRESSED</a>
    <a href="#">COLOR SHOP</a>
    <a href="#">$29.99 & UNDER JEANS</a>
    <a href="#">DENIM SHORTS</a>
    <a href="#">SALE JEANS</a>
  </div>
  </div>
 <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=shoes">SHOES</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">BOOTS $20 AND UNDER</a>
    <a href="#">SALE SHOES</a>
    <a href="#">HEELS</a>
    <a href="#">BOOTS & BOOTIES</a>
    <a href="#">SNEAKERS</a>
    <a href="#">SANDALS</a>
    <a href="#">FLATS</a>
    <a href="#">GLASS COLLECTION</a>
  </div>
</div>
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=tops">TOPS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">FASHION KNIT TOPS</a>
    <a href="#">SHIRT & BLOUSES</a>
    <a href="#">BODYSUITS</a>
    <a href="#">SWEATERS</a>
    <a href="#">JACKETS</a>
    <a href="#">EVERYDAY TOPS</a>
    <a href="#">GOING OUT TOPS</a>
    <a href="#">SALE TOPS</a>
    <a href="#">SALE SWEATERS & JACKETS</a>
    <a href="#">SALE $10 & UNDER</a>
  </div>
</div>
 
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=bottoms">BOTTOMS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">PANTS</a>
    <a href="#">LEGGINGS</a>
    <a href="#">SKIRTS</a>
    <a href="#">SHORTS</a>
    <a href="#">DENIM SHORTS</a>
    <a href="#">BIKER SHORTS</a>
    <a href="#">SALE BOTTOMS</a>
    <a href="#">SALE SHORTS</a>
  </div>
</div>

  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=jackets">JACKETS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">SALE JACKETS</a>
    <a href="#">BLAZERS</a>
    <a href="#">DENIM</a>
    <a href="#">CASUAL</a>
    <a href="#">LEATHER + SUEDE</a>
    <a href="#">FAUX FUR</a>
    <a href="#">OUTERWEAR</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=sweaters">SWEATERS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">SALE SWEATERS</a>
    <a href="#">CARTIGANS & DUSTERS</a>
    <a href="#">SWEATER DRESSES</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=luxe">LUXE</a>
  <div class="dropdown-content">
    <a href="#">DRESSES</a>
    <a href="#">MATCHING SETS</a>
    <a href="#">ROMPERS + JUMPSUITS</a>
    <a href="#">CURVE + PLUS</a>
  </div>
</div>

  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=matching">MATCHING SETS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">PANT SETS</a>
    <a href="#">SKIRT SETS</a>
    <a href="#">SHORT SET</a>
    <a href="#">SALE MATCHING SET</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=rompers">ROMPERS AND JUMPSUITS</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">ROMPERS</a>
    <a href="#">JUMPSUITS</a>
    <a href="#">SALE ROMPERS AND JUMPSUITS</a>
  </div>
</div>

  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=active">ACTIVEWEAR</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">ACTIVE TOPS</a>
    <a href="#">ACTIVE BOTTOMS</a>
    <a href="#">ACTIVE SETS</a>
    <a href="#">GRAPHIC TOPS</a>
    <a href="#">SALE ACTIVEWEAR</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=swim">SWIM</a>
  <div class="dropdown-content">
    <a href="#">ONE-PIECE</a>
    <a href="#">TWO-PIECE</a>
    <a href="#">COVER-UPS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=accessories">ACCESSORIES</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">HANDBAGS</a>
    <a href="#">SUNGLASSES</a>
    <a href="#">BELTS</a>
    <a href="#">HATS & SCARVES</a>
    <a href="#">LEGWEAR</a>
    <a href="#">HAIR ACCESSORIES</a>
    <a href="#">BEAUTY</a>
    <a href="#">JEWELRY</a>
    <a href="#">EARRINGS</a>
    <a href="#">NECKLACES</a>
    <a href="#">BRACELETS</a>
    <a href="#">BODY JEWELRY</a>
    <a href="#">RINGS</a>
    <a href="#">SALE ACCESSORIES</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=lingerie">LINGERIE & INTIMATES</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">SEXY LINGERIE</a>
    <a href="#">TEDDIES + BABYDOLLS</a>
    <a href="#">LINGERIE SETS</a>
    <a href="#">STOCKINGS</a>
    <a href="#">LINGERIE ESSENTIALS</a>
    <a href="#">BRAS + BRALETTES</a>
    <a href="#">PANTIES</a>
    <a href="#">SLEEPWEAR + ROBES</a>
    <a href="#">SHAPEWEAR</a>
    <a href="#">SALE LINGERIE</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=wear">WEAR TO WORK</a>
  <div class="dropdown-content">
    <a href="#">MATCHING SETS</a>
    <a href="#">DRESSES</a>
    <a href="#">TOPS</a>
    <a href="#">JACKETS</a>
    <a href="#">BOTTOMS</a>
    <a href="#">SCRUBS</a>
    <a href="#">PLUS</a>
  </div>
</div>

  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=collections">COLLECTIONS</a>
  <div class="dropdown-content">
    <a href="#">BAD GIRLS ONLY</a>
    <a href="#">MARIGOLD COLLECTION</a>
    <a href="#">THE DENIM SHOP</a>
    <a href="#">NEON COLLECTION</a>
    <a href="#">SHOP CARDI</a>
    <a href="#">KYLIE</a>
    <a href="#">BACK IN STOCK</a>
    <a href="#">BEST SELLERS</a>
    <a href="#">COMING SOON</a>
    <a href="#">ESSENTIALS</a>
    <a href="#">SCRUBS</a>
    <a href="#">UNISEX</a>
    <a href="#">HOT MAMA</a>
    <a href="#">NOVABABE EXPERIENCE</a>
    <a href="#">SHOP INSTAGRAM</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=plus">PLUS SIZE & CURVE</a>
  <div class="dropdown-content">
    <a href="#">NEW ARRIVALS</a>
    <a href="#">MARIGOLD COLLECTION</a>
    <a href="#">NEON COLLECTION</a>
    <a href="#">DRESSES</a>
    <a href="#">BLACK DRESSES</a>
    <a href="#">SPECIAL OCCASION DRESSES</a>
    <a href="#">JEANS</a>
    <a href="#">HIGH RISE JEANS</a>
    <a href="#">LEGGINGS</a>
    <a href="#">BOTTOMS</a>
    <a href="#">LINGERIE</a>
    <a href="#">TOPS</a>
    <a href="#">GRAPHIC TOPS</a>
    <a href="#">JACKETS</a>
    <a href="#">SWEATERS</a>
    <a href="#">MATCHING SETS</a>
    <a href="#">ROMPERS AND JUMPSUITS</a>
    <a href="#">ACTIVEWEAR</a>
    <a href="#">SWIMWEAR</a>
    <a href="#">PLUS SALE DRESSES</a>
    <a href="#">PLUS SALE SWEATERS & JACKETS</a>
    <a href="#">PLUS SALE JEANS</a>
    <a href="#">$10 & UNDER SALE</a>
    <a href="#">$15 & UNDER SALE</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=women&cat=sale">SALE</a>
  <div class="dropdown-content">
    <a href="#">LAST CHANCE</a>
    <a href="#">BLOWOUT</a>
    <a href="#">$5 AND UNDER</a>
    <a href="#">$10 AND UNDER</a>
    <a href="#">$15 AND UNDER</a>
    <a href="#">$20 AND UNDER</a>
    <a href="#">SALE SWEATERS & JACKETS</a>
    <a href="#">SALE TOPS $10 & UNDER</a>
    <a href="#">BOOTS $20 & UNDER</a>
    <a href="#">SALE DRESSES</a>
    <a href="#">SALE JEANS</a>
    <a href="#">SALE TOPS</a>
    <a href="#">SALE SWEATERS</a>
    <a href="#">SALE JACKETS</a>
    <a href="#">SALE BOTTOMS</a>
    <a href="#">SALE SHOES</a>
    <a href="#">SALE ROMPERS AND JUMPSUITS</a>
    <a href="#">SALE MATCHING SETS</a>
    <a href="#">SALE LINGERIE</a>
    <a href="#">SALE ACCESSORIES</a>
    <a href="#">SALE BEAUTY</a>
  </div>
</div>
  <img src="/graphics/images/downs.png"> 

</div>

<div id="Paris" class="tabcontent" style="text-align:center;position:absolute;left:0px;background-color:black;color:white;padding:10px 10px 10px 10px">
  <a href="categories.asp?s=men&cat=all">SHOP ALL</a> <img src="/graphics/images/downs.png"> 
    <div class="dropdown">
  <a class="dropbtn" href="newarrivals.asp?s=men">NEW ARRIVALS</a>
  <div class="dropdown-content">
    <a href="#">COMING SOON</a>
    <a href="#">NEO COLLECTION</a>
    <a href="#">NEW GRAPHICS</a>
    <a href="#">NEW TOPS</a>
    <a href="#">NEW DENIMS</a>
    <a href="#">NEW BOTTOMS</a>
  </div>
</div>

  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=plus">PLUS+</a>
  <div class="dropdown-content">
    <a href="#">TOPS</a>
    <a href="#">BOTTOMS</a>
    <a href="#">DENIM</a>
    <a href="#">ACTIVE</a>
    <a href="#">GRAPHICS</a>
    <a href="#">COMING SOON</a>
  </div>
</div>
 
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=jeans">JEANS</a>
  <div class="dropdown-content">
    <a href="#">SKINNY</a>
    <a href="#">SLIM</a>
    <a href="#">MOTO</a>
    <a href="#">JEAN JACKETS</a>
  </div>
</div>
 
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=bottoms">BOTTOMS</a>
  <div class="dropdown-content">
    <a href="#">PANTS</a>
    <a href="#">SHORTS</a>
    <a href="#">CARGO BOTTOMS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=active">ACTIVEWEAR</a>
  <div class="dropdown-content">
    <a href="#">JOGGERS</a>
    <a href="#">JACKETS & HOODIES</a>
    <a href="#">TRACKSUITS</a>
  </div>
</div>
 
  <img src="/graphics/images/downs.png"> 
  <a href="categories.asp?s=men&cat=active">HOODIES</a> <img src="/graphics/images/downs.png"> 
  <a href="categories.asp?s=men&cat=active">TANK TOPS</a> <img src="/graphics/images/downs.png"> 
  <a href="categories.asp?s=men&cat=sweater">SWEATERS</a> <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=graphics">GRAPHICS</a>
  <div class="dropdown-content">
    <a href="#">SHORT SLEEVE</a>
    <a href="#">LONG SLEEVE</a>
    <a href="#">HOODIE & JOGGERS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=tops">TOPS</a>
  <div class="dropdown-content">
    <a href="#">SWEATERS & SWEATSHIRTS</a>
    <a href="#">POLOS</a>
    <a href="#">SHIRTS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=basics">BASICS</a>
  <div class="dropdown-content">
    <a href="#">BASIC TOPS</a>
    <a href="#">BASIC BOTTOMS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
  <a href="categories.asp?s=men&cat=jacket">JACKET & OUTERWEAR</a> <img src="/graphics/images/downs.png"> 
  <a href="categories.asp?s=men&cat=accessories">ACCESSORIES</a> <img src="/graphics/images/downs.png">
  <a href="categories.asp?s=men&cat=shoes">SHOES</a> <img src="/graphics/images/downs.png">
  <div class="dropdown">
  <a class="dropbtn" href="categories.asp?s=men&cat=sale">SALE</a>
  <div class="dropdown-content">
    <a href="#">BLOWOUT</a>
    <a href="#">$10 AND UNDER</a>
    <a href="#">$15 AND UNDER</a>
    <a href="#">$20 AND UNDER</a>
    <a href="#">SALE DENIM</a>
    <a href="#">SALE TOPS</a>
    <a href="#">SALE BOTTOMS</a>
    <a href="#">SALE ACTIVEWEAR</a>
    <a href="#">SALE GRAPHICS</a>
  </div>
</div>
  
  <img src="/graphics/images/downs.png"> 
</div>

<script>
function openCity(evt, cityName) {
  var i, tabcontent, tablinks;
  tabcontent = document.getElementsByClassName("tabcontent");
  for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
  }
  tablinks = document.getElementsByClassName("tablinks");
  for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
  }
  document.getElementById(cityName).style.display = "block";
  evt.currentTarget.className += " active";
  if(document.getElementById("London").style.display != 'none') {
  	document.getElementById("defaultOpen").style.backgroundColor = "#000";
  	document.getElementById("men").style.backgroundColor = "#eee";
	  document.getElementById("defaultOpen").style.color = "#fff";
  	document.getElementById("men").style.color = "#000";
    document.getElementById("two").style.display = "";
    document.getElementById("one").style.display = "none";
    <%
    if instr(Request.ServerVariables("http_user_agent"),"index.asp") > -1 then
    %>
    document.getElementById("womenw").style.display = '';
    document.getElementById("notWomen").style.display = 'none';
    document.getElementById("thisiswomen").style.display = '';
    document.getElementById("thisismen").style.display = 'none';
    <%
    end if
    %>
  } else if(document.getElementById("men").style.display != 'none') {
  	document.getElementById("defaultOpen").style.backgroundColor = "#eee";
  	document.getElementById("men").style.backgroundColor = "#000";
	  document.getElementById("defaultOpen").style.color = "#000";
  	document.getElementById("men").style.color = "#fff";
    document.getElementById("one").style.display = "";
    document.getElementById("two").style.display = "none";
    <%
    if instr(Request.ServerVariables("http_user_agent"),"index.asp") > -1 then
    %>
    document.getElementById("womenw").style.display = 'none';
    document.getElementById("notWomen").style.display = '';
    document.getElementById("thisiswomen").style.display = 'none';
    document.getElementById("thisismen").style.display = '';
    <%
    end if
    %>
  }
}

<%
dim s

s = request("s")

if s = "" then
  s = "women"
end if

if s = "women" then
%>

// Get the element with id="defaultOpen" and click on it
document.getElementById("defaultOpen").click();

<%
else
%>

// Get the element with id="defaultOpen" and click on it
document.getElementById("men").click();

<%
end if
%>
</script>

   <button onclick="topFunction()" id="myBtn" title="Go to top">Top</button>