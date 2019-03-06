<%
    dim file, form

    set form = Server.CreateObject("MyRequest.Form")
    form.Init


    dim sqlstr


dim Conn
Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

dim up, itemid, newarrival, name, price, description, category, image, weight, discount, women, men

itemid = form("itemid")

up = form("up")

  if up = "yes" then
    
    dim oRS, bestseller

    newarrival = form("newarrival")

    if newarrival = "on" then

      newarrival = "1"

    else

      newarrival = "NULL"

    end if

    bestseller = form("bestseller")

    if bestseller = "on" then

      bestseller = "1"

    else

      bestseller = "NULL"

    end if

    set file = form.Item("file").Item(1)

    Response.Write("FileName =" & file.FileName & "<br>")
    Response.Write("FileExt  =" & file.FileExt & "<br>")
    Response.Write("FilePath =" & file.FilePath & "<br>")
    Response.Write("MimeType =" & file.ContentType & "<br><br>")
    'Response.Write("FileSize =" + file.TotalBytes)

    dim fileStream

    set fileStream = server.CreateObject("ADODB.Stream")

    fileStream.Open()

    fileStream.Type = 1
     'adTypeBinary

    fileStream.Write(file.Value)

    fileStream.SaveToFile Server.MapPath("productitems") + "\\" + file.FileName, 2
    'adSaveCreateOverWrite

    fileStream.Close()

    name = form("name")

    price = form("price")

    weight = form("weight")

    description = form("dsc")

    category = form("category")

    women = form("women")

    men = form("men")

    discount = form("discount")

    dim theitemid
    theitemid = form("theitemid")

    if theitemid = "" then
      theitemid = "NULL"
    end if

    if women = "on" and men <> "on" then

      sqlstr = "UPDATE products2 set women = 1, weight = "&weight&", bestseller = "&bestseller&", recommendation = "&theitemid&", newarrivals = "&newarrival&", name='"&name&"',price="&price&",dsc='"&description&"',category='"&category&"',image='"&file.FileName&"', discount=25 WHERE itemid = " & itemid

    elseif women = "on" and men = "on" then

      sqlstr = "UPDATE products2 set women = 1, men = 1, weight = "&weight&", bestseller = "&bestseller&", recommendation = "&theitemid&", newarrivals = "&newarrival&", name='"&name&"',price="&price&",dsc='"&description&"',category='"&category&"',image='"&file.FileName&"', discount=25 WHERE itemid = " & itemid

    elseif women <> "on" and men = "on" then

      sqlstr = "UPDATE products2 set men = 1, weight = "&weight&", bestseller = "&bestseller&", recommendation = "&theitemid&", newarrivals = "&newarrival&", name='"&name&"',price="&price&",dsc='"&description&"',category='"&category&"',image='"&file.FileName&"', discount=25 WHERE itemid = " & itemid
    
    end if
'response.write sqlstr
    Conn.Execute(sqlstr)

end if

Conn.close()
%>