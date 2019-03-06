<!--#include file="inc/header.inc"-->
<div style="width:263px;height:487px;float:left;position:relative;top:62px;left:-240px;background-color:#333;">
<!--#include file="inc/sidemenu.inc"-->
</div>
<div id="one" style="display:none;float:left;position:relative;top:-425px;left:263px;text-align:center;width:1000px">
  <img src="/graphics/images/blowout80per.jpg">
</div>

<div id="two" style="position:relative;top:-425px;left:263px;float:left;text-align:center;width:1000px">
  <img src="/graphics/images/blowout80per2.png">
</div>

    <div id="thisiswomen" style="text-align:center;position:absolute;top:660px;left:0px;width:100%;height:60px;background-color:#e8e8e8">
    <h1>Shop Below</h1>
    <h1>Featured</h1>
    
<%
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2 where women = 1 and featured = 1    ORDER BY DBMS_RANDOM.VALUE) where rownum between 1 and 4;"

    set oRS = Conn.Execute(ssql)

    count = 0

    dim f

    f = false

    while not oRS.eof
    if count < 4 then
    %>
    <div style = "float:left; width: 300px; height: 480px;color:black">
      <div style="text-align:center">
        <a style="color:black" href='items.asp?itemid=<%=oRS("itemid")%>'><%=oRS("name")%></a><br>
      </div>
      <div style="text-align:center">
        <a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:black"><nobr><form method="post" action="cart.asp"><input type="hidden" name="itemid" value=<%=oRS("itemid")%>><img src="/graphics/images/view.png" onmouseover="this.src='/graphics/images/view-mo.png'" onmouseout="this.src='/graphics/images/view.png'"><input type="image" src="/graphics/images/add.png" onmouseover="this.src='/graphics/images/add-mo.png'" onmouseout="this.src='/graphics/images/add.png'" border="0" name="submit"></form></nobr><br><img src='/productitems/<%=oRS("image")%>' width="80" height="230"></a><br>
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
      end if
      count = count + 1
      oRS.movenext
    wend
%>
</div>
    <div id="thisismen" style="display:none;text-align:center;position:absolute;top:660px;left:0px;width:100%;height:60px;background-color:#e8e8e8">
    <h1>Shop Below</h1>
    <h1>Featured</h1>
    
<%
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2 where men = 1 and featured = 1    ORDER BY DBMS_RANDOM.VALUE) where rownum between 1 and 4;"

    set oRS = Conn.Execute(ssql)

    count = 0

    while not oRS.eof
    if count < 4 then
    %>
    <div style = "float:left; width: 300px; height: 480px;color:black">
      <div style="text-align:center">
        <a style="color:black" href='items.asp?itemid=<%=oRS("itemid")%>'><%=oRS("name")%></a><br>
      </div>
      <div style="text-align:center">
        <a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:black"><nobr><form method="post" action="cart.asp"><input type="hidden" name="itemid" value=<%=oRS("itemid")%>><img src="/graphics/images/view.png" onmouseover="this.src='/graphics/images/view-mo.png'" onmouseout="this.src='/graphics/images/view.png'"><input type="image" src="/graphics/images/add.png" onmouseover="this.src='/graphics/images/add-mo.png'" onmouseout="this.src='/graphics/images/add.png'" border="0" name="submit"></form></nobr><br><img src='/productitems/<%=oRS("image")%>' width="80" height="230"></a><br>
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
      end if
      count = count + 1
      oRS.movenext
    wend
%>
</div>
<div id="womenw" style="float:left;width:100%;z-index:100;position:relative;top:50px">
<%
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2 where women = 1 and featured is null   ORDER BY DBMS_RANDOM.VALUE);"

    set oRS = Conn.Execute(ssql)

    count = 0

    f = false

    while not oRS.eof

    if count mod 4 = 0 then
    %>
      <div style="float:left;width:100%;height:350px">
    <%
    end if
    %>
    <div style = "float:left; width: 300px; height: 350px;color:black">
      <div style="text-align:center">
        <a style="color:black" href='item.asp?itemid=<%=oRS("itemid")%>'><%=oRS("name")%></a><br>
      </div>
      <div style="text-align:center">
        <a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:black"><nobr><form method="post" action="cart.asp"><input type="hidden" name="itemid" value=<%=oRS("itemid")%>><img src="/graphics/images/view.png" onmouseover="this.src='/graphics/images/view-mo.png'" onmouseout="this.src='/graphics/images/view.png'"><input type="image" src="/graphics/images/add.png" onmouseover="this.src='/graphics/images/add-mo.png'" onmouseout="this.src='/graphics/images/add.png'" border="0" name="submit"></form></nobr><img src='/productitems/<%=oRS("image")%>' width="120" height="180"></a><br><br>
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
      f = true
    else
      f = false
    end if
    %>
<%
  oRS.movenext
  count = count + 1
wend

if not f then
%>
</div>
<%
end if

    %>
</div>
<div id="notWomen" style="display:none;float:left;width:100%;z-index:101;position:relative;top:50px">
<%
  Dim ssql
  ssql = "SELECT image, discount, itemid, name, price FROM   (    SELECT *    FROM   products2 where men = 1 and women is null and men is not null and featured is null   ORDER BY DBMS_RANDOM.VALUE);"

    set oRS = Nothing
    set oRS = Conn.Execute(ssql)

    count = 0

    f = false

    while not oRS.eof
    if count mod 4 = 0 then
    %>
      <div style = "float:left;width:100%;height:350px">
    <%
    end if
    %>
    <div style = "float:left; width: 300px; height: 350px;color:black">
      <div style="text-align:center">
        <a style="color:black" href='item.asp?itemid=<%=oRS("itemid")%>'><%=oRS("name")%></a><br>
      </div>
      <div style="text-align:center">
        <a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:black"><nobr><form method="post" action="cart.asp"><input type="hidden" name="itemid" value=<%=oRS("itemid")%>><img src="/graphics/images/view.png" onmouseover="this.src='/graphics/images/view-mo.png'" onmouseout="this.src='/graphics/images/view.png'"><input type="image" src="/graphics/images/add.png" onmouseover="this.src='/graphics/images/add-mo.png'" onmouseout="this.src='/graphics/images/add.png'" border="0" name="submit"></form></nobr><img src='/productitems/<%=oRS("image")%>' width="120" height="180"></a><br><br>
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
      f = true
    else
      f = false
    end if
    %>
<%
  oRS.movenext
  count = count + 1
wend

if not f then
%>
</div>
<%
end if

    %>
  </div>
    <%

Conn.close()
%>
<h1>&nbsp;</h1>
<!--#include file="inc/footer.inc"-->