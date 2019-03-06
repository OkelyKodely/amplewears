<%
    dim file, form

    set form = Server.CreateObject("MyRequest.Form")
    form.Init


    dim sqlstr

dim Conn
Set Conn = Server.Createobject("ADODB.Connection")
Conn.open "Driver={Oracle in XE};Server=localhost; Uid=sa;Pwd=coppersink21;"

dim up, newarrival, name, price, description, category, image, weight, women, men, discount, featured

up = form("up")

  if up = "yes" then
    
    dim bestseller

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

    featured = form("featured")

    discount = form("discount")

    dim theitemid
    theitemid = form("theitemid")

    if theitemid = "" then
      theitemid = "NULL"
    end if

    if featured = "on" then

    if women = "on" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women, men, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1, 1)"
    end if

    if women = "on" and men = "off" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1)"
    end if

    if women = "on" and men = "" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1)"
    end if

    if women = "off" and men = "off" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    if women = "" and men = "" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    if women = "off" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, men, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1)"
    end if

    if women = "" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, men, featured) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1)"
    end if

    else

    if women = "on" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women, men) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1, 1)"
    end if

    if women = "on" and men = "off" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    if women = "on" and men = "" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, women) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    if women = "off" and men = "off" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&")"
    end if

    if women = "" and men = "" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&")"
    end if

    if women = "off" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, men) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    if women = "" and men = "on" then
    sqlstr = "INSERT INTO products2 (weight, bestseller, recommendation, newarrivals, name, price, dsc, category, image, discount, men) VALUES ("&weight&","&bestseller&","&theitemid&","&newarrival&",'"&name&"',"&price&",'"&description&"','"&category&"','"&file.FileName&"', "&discount&", 1)"
    end if

    end if

response.write sqlstr
    Conn.Execute(sqlstr)
  end if

  Conn.Close()
%>