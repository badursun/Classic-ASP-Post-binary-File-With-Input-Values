# Classic-ASP-Post-binary-File-With-Input-Values
How to read file binary and post to remote server and bonus with inputs and values.

# Is it possible?
Yes! Only use some MSMXL object, FileSystemObject, ADODB.Stream and Magic.

# How To
1-) ADODB.Stream is a first object you can use to read/write text and binary files. The object is included in ADO 2.5 and later.
2-) ReadBinaryFile OR ReadTextFile AND SaveBinaryData OR SaveTextData on memory.
3-) Convert binary data to Text for submiting.
4-) Build WebKitFormBoundaryd (this is very important)
4-) Post data to Remote URL with Binary data in WebKitFormBoundaryd

# What the WebKitFormBoundaryd
Each item in a multipart message is separated by a boundary marker. Webkit based browsers put "WebKitFormBoundary" in the name of that boundary. 

# What does the random string after a WebKitFormBoundary do\mean?
That's just the typical way of how a so called "boundary" between different parts of a mime structure is defined. The receiving side can tell the different parts apart by this. Same logic is used in different things, email messages too for example. Actually it is not the "random part" of that boundary that counts. The whole string is matched. It is simply a convention that each software uses a unique prefix string for such boundaries for transparency reasons. But in general the only requirement is that the chosen string must be unique throughout all contained data. Unique obviously except for the corresponding boundaries which must use exactly the same string.

# How To User Script
Just include file or class to your script

  <!--#include file="/yourPath/FPost.asp"-->

And set class

  <%
    Dim FPost 
    Set FPost = New FilePost 
  %>

And add some file and inputs

  <%
      ' Add Some File
      FPost.AddFile("1.png")
      FPost.AddFile("") ' Return ERROR
      FPost.AddFile("2.jpg") ' Return ERROR because file not exist
      FPost.AddFile("sample.png")
      FPost.AddFile("test.txt")
      FPost.AddFile("3.png")
      FPost.AddFile("3png.zip")

      ' Add Some Data (inputName, inputValue)
      FPost.AddText "name", "Anthony Burak"
      FPost.AddText "email", "badursun@gmail.com"
      FPost.AddText "surname", "DURSUN"

      ' Post File URL And Post Type
      FPost.PostURL("http://remote_url/maybe.php")
      FPost.FormType("POST") ' POST, PUT, DELETE

      ' Get Post HTTPStatus Answer (Default:200=Success)
      If FPost.PostFiles() = True Then 
        Response.Write "SUCCESS"&vbcrlf
        Response.Write "STATUS CODE: "&FPost.HTTPStatus()
        Response.Write "HTTP ANSWER: "&FPost.HTTPAnswer()
      Else
        Response.Write "FAILED"&vbcrlf
        Response.Write "STATUS CODE: "&FPost.HTTPStatus()
        Response.Write "HTTP ANSWER: "&FPost.HTTPAnswer()
      End If
  Set FPost = Nothing
  %>





