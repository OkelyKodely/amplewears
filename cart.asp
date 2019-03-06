<!--#include file="inc/header.inc"-->
<style>
    .title {
      font-size: 200%;
      font-stretch: condensed;
      font-weight: bold;
    }
  </style>
  
    <div style="width:1000px;float:left">

    <div style="text-align:left;float:left;font-size:40px;color:white;background-color:gray;position:relative;top:90px;left:-240px;width:760px;height:70px">
      <!--<img src="/graphics/images/shopping-cart-banner.png">-->
      <br>
      &nbsp;YOUR CART
    </div>

    </div>

    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>
    <h1>&nbsp;</h1>

    <div style="width:1200px;float:left;position:relative;left:10px">
    <%
    dim itemid

    itemid = request.form("itemid")

    dim sqlstr

    dim cartid

    dim clearcart

    dim w

    cartid = session("cartid")
    'response.write cartid

    clearcart = request.querystring("do")

    if clearcart = "clear" then

      sqlstr = "DELETE FROM cart2 WHERE cid = " & cartid

      Conn.Execute(sqlstr)

    end if

    if itemid <> "" then

      if isnull(cartid) or cartid = "" then

        sqlstr = "SELECT cid AS id FROM cart2 order by cid desc"
'response.write sqlstr
        set rs = Conn.Execute(sqlstr)

        if not rs.eof then
          cartid = cint(rs("id")) + 1
        'response.write "1wer" &cartid
        else
          cartid = "1"
        'response.write "wer" & cartid
        end if

        session("cartid") = cartid

        'response.write "asdf " & cartid
        
      end if

      sqlstr = "INSERT INTO cart2 (cid,itemid) VALUES ("&cartid&","&itemid&")"

      'response.write sqlstr
   
     Conn.Execute(sqlstr)
 
   end if

   sqlstr = "SELECT * FROM cart2 WHERE cid=" & cartid

   if cartid = "" then
    set oRS = nothing
   else   
    set oRS = Conn.Execute(sqlstr)

   end if

   dim total

   if cartid <> "" then
   %><br><br>
    <div style="float:left;width:590px;height:100%">


    <%
if oRS.eof then
%>
<h1>There are no items in your cart.</h1>
<%
end if
w = 0
    while not oRS.eof
      dim iid
      iid = oRS("itemid")

      sqlstr = "SELECT weight, image, name, category, discount, price, dsc FROM products2 WHERE itemid = " & iid

      set r_s = Conn.Execute(sqlstr)

      while not r_s.eof
        %>
    <div style = "width:900px;float:left">
    <div style="float:left;width:60px">image</div>
    <div style="float:left;width:260px">item id</div>
    <div style="float:left;width:110px">category</div>
    <div style="float:left;width:230px">price</div>
    </div>


      <div style = "width:750px;height:170px;border:1px solid gray">

      <div style="float:left;width:60px"><div style="height:40px;width:40px"><img src=/productitems/<%=r_s("image")%>></div>&nbsp;</div>&nbsp;
      <div style="float:left;width:260px"><a style="color:black" href='items.asp?itemid=<%=iid%>'><%=r_s("name")%></a></div>&nbsp;
      <div style="float:left;width:110px"><%=r_s("category")%></div>&nbsp;
      <div style="float:left;width:230px">$<%response.write formatnumber((100-cdbl(r_s("discount")))*cdbl(r_s("price"))/100,2)%>
<br>
<b>Description:</b> <%=left(r_s("dsc"),150)%>...
      </div>

      </div><br>
        <%
        total = total + (100-cdbl(r_s("discount")))*cdbl(r_s("price"))/100
        w = w + cdbl(r_s("weight"))
        r_s.movenext
      wend

      oRS.movenext
    wend
 
  %>
    </div>
  <%
  end if

  if total > 0 then
  %>

<div style="border:10px solid gray;float:left;width:262px;height:300px;position:relative;left:200px;top:0px">
  <h3>Summary</h3>
<h4 style="color:red">Subtotal: $<%=formatnumber(total,2)%></h4>
<a href="cart.asp?do=clear" style="color:#ff0000"><button>Clear cart</button></a>
<script>
  var x = document.getElementById("mi");
  x.innerHTML = '<img src="/graphics/images/shipping.png"><font size=-2>$20 flat fee fore shipping using UPS Ground Shipping method.  Shipping takes usually 1 ~ 6 business days.</font>';
  x.style.display = "block";
</script>
<br><br>
  <form method="post" action="buy.asp">
    <input type="hidden" name="t" value=<%=total%>>
    <input type="hidden" name="w" value=<%=w%>>
    <input type="submit" value="Check Out" style="background-color:gray;color:white;font-size:20px;font-weight:bold;width:200px; height:70px">
  </form>

</div>
<%
end if
%>

<%
Conn.close()
%>

  </div>

<!--#include file="inc/footer.inc"-->