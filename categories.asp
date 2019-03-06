<!--#include file="inc/header.inc"-->
<div style="width:263px;height:483px;float:left;position:relative;top:62px;left:-240px;background-color:#333;">
<!--#include file="inc/sidemenu.inc"-->
</div>
<%
Dim a, b
a = request("a")
b = request("b")

if a = "" then
  a = 1
  b = 16
end if

  Dim ssql
  if s = "men" then
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2  WHERE category = '"&cat&"' and men = 1) WHERE rownum between "&a&" AND "&b&";"
  else
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2  WHERE category = '"&cat&"' and women = 1) WHERE rownum between "&a&" AND "&b&";"
  end if
    set oRS = Conn.Execute(ssql)

    count = 0

    while not oRS.eof

    if count mod 4 = 0 then
    %>
    <br><br><br><br>
      <div style="order:1px solid black;z-index:2;text-align:left;width:100%;height:600px">
    <%
    end if
    %>
    <div style = "position:relative;left:-200px;boder:1px solid black;float:left; width:250px; height: 480px;color:black">
      <div style="text-align:center">
        <a style="color:black" href='items.asp?itemid=<%=oRS("itemid")%>'><%=oRS("name")%></a><br>
      </div>
      <div style="text-align:center">
        <a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:black"><nobr><form method="post" action="cart.asp"><input type="hidden" name="itemid" value=<%=oRS("itemid")%>><img src="/graphics/images/view.png" onmouseover="this.src='/graphics/images/view-mo.png'" onmouseout="this.src='/graphics/images/view.png'"><input type="image" src="/graphics/images/add.png" onmouseover="this.src='/graphics/images/add-mo.png'" onmouseout="this.src='/graphics/images/add.png'" border="0" name="submit"></form></nobr><img src='/productitems/<%=oRS("image")%>' width="200" height="230"></a><br><br>
        <%
        if cint(oRS("discount")) <> 0 then
        %>
                <a href='items.asp?itemid=<%=oRS("itemid")%>'><font color="red" size="4">On Sale: $<%response.write formatnumber((100-cdbl(oRS("discount")))*cdbl(oRS("price"))/100,2)%></font></a><br>
        <%
        else
        %>
                <a href='items.asp?itemid=<%=oRS("itemid")%>'><font color="#cc0000" size="4">Price: $<%response.write formatnumber((100-cdbl(oRS("discount")))*cdbl(oRS("price"))/100,2)%></font></a><br>
        <%
        end if
        %>
      </div>
    </div>
    <%
    if count mod 4 = 3 then
    %>
      </div>
    <%
    end if
    %>
<%
  oRS.movenext
  count = count + 1
wend

    %>
      </div>
    <%

Conn.close()
%>
<div style="height:700px;width:100%">&nbsp;</div>
<!--#include file="inc/footer.inc"-->