<!--#include file="inc/header.inc"-->
  <div style="position:relative;top:100px">
  
    <div class="shoptitle">
      <span style="font-size:26">Your Search Results</span>
    </div>

    <div style="width:650px">
    <%
    dim sqlstr

    dim query

    query = request("search")

    query = replace(query, " ", "%")

    dim a, b

    a = request("a")

    b = request("b")

    if a = "" then
    a = 1
    b = 6
    end if

    dim ssql

    ssql = "SELECT * FROM   (    SELECT *    FROM   products2 where name like '%"&query&"%') WHERE rownum between "&a&" AND "&b&";"

    set oRS = Conn.Execute(ssql)

    %>
    
<Br><br>

    <div style = "width:100%;">

    <div style="float:left;width:840px">Results for: <%=query%><hr style="width:1200px">
</div>
    </div>
      <%
      dim count

      count = 0

    while not oRS.eof and query <> ""
    if count = 0 then
    %>
      <div stylef = "width:800px;float:left;position:relative;left:-100px">
      <div style = "float:left; width: 50px; height: 30px">
      <a href='shopresults.asp?query=<%=query%>&a=<%=a-6%>&b=<%=b-6%>'>less...</a>
    </div>
    <%
    end if

      %>
      <div style = "width:1000px;float:left">

      <div style="width:220px"><%=oRS("itemid")%> &nbsp;&nbsp;&nbsp;&nbsp;<a href='items.asp?itemid=<%=oRS("itemid")%>' style="color:red"><%=oRS("name")%></a></div>
      <div style="width:220px">$<%=oRS("price")%> USD</div>
      <div style="width:220px">Category: <%=oRS("category")%></div>
      <div style="width:220px"><img src='/productitems/<%=oRS("image")%>'></div>
      <div style="width:100%">Description: <%=oRS("dsc")%></div>

      </div>

      <h1>&nbsp;</h1>
      <%
      oRS.movenext
      count = count + 1
    wend

Conn.close()
%>
<!--#include file="inc/footer.inc"-->