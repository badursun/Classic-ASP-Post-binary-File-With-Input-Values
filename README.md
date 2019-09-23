# Introduction
Read binary file and POST Remote SErver with input and values on Classic ASP

FPost Class is helper for read binary data and post to remote server with WebKitFormBoundary standarts. Just define the class and specify the physical files you want to add. Automatically checks for the presence of files. Makes the related MimeType definitions.

# Usage

## How To User Script
Just include file or class to your script
```asp
  <!--#include file="/yourPath/FPost.asp"-->
```

And set class

```asp
<%
	Dim FPost 
	Set FPost = New FilePost 
%>
```

And add some file and inputs

```asp
  <%
    ' Add Some File
    FPost.AddFile("test_files/1.png")
    FPost.AddFile("test_files/test.txt")
    FPost.AddFile("test_files/3.png")
    FPost.AddFile("test_files/3png.zip")

    ' Add Some Not Exist File
    FPost.AddFile("")
    FPost.AddFile("test_files/2.jpg")
    FPost.AddFile("test_files/sample.png")

    ' Add Some Data (inputName, inputValue)
    FPost.AddText "adi", "Anthony Burak"
    FPost.AddText "eposta", "badursun@gmail.com"
    FPost.AddText "soyadi", "DURSUN"
    
    ' Post File URL And Post Type
    FPost.PostURL("http://demoadresi.com/GET_FILE/")
    FPost.FormType("POST")

    ' Get Post HTTPStatus Answer (Default:200=Success)
    If FPost.PostFiles() = True Then 
      Response.Write "SUCCESS"&vbcrlf
      Response.Write "STATUS CODE: "&FPost.HTTPStatus()
      Response.Write "HTTP ANSWER: "&FPost.HTTPAnswer()
    Else
      Response.Write "SUCCESS"&vbcrlf
      Response.Write "STATUS CODE: "&FPost.HTTPStatus()
      Response.Write "HTTP ANSWER: "&FPost.HTTPAnswer()
    End If
    Set FPost = Nothing
  %>
```

# Some Information 

## Read binary file and POST Remote SErver with input and values on Classic ASP
How to read file binary and post to remote server and bonus with inputs and values.

## Is it possible?
Yes! Only use some MSMXL object, FileSystemObject, ADODB.Stream and Magic.

## How To
1-) ADODB.Stream is a first object you can use to read/write text and binary files. The object is included in ADO 2.5 and later.
2-) ReadBinaryFile OR ReadTextFile AND SaveBinaryData OR SaveTextData on memory.
3-) Convert binary data to Text for submiting.
4-) Build WebKitFormBoundaryd (this is very important)
4-) Post data to Remote URL with Binary data in WebKitFormBoundaryd

## What the WebKitFormBoundary
Each item in a multipart message is separated by a boundary marker. Webkit based browsers put "WebKitFormBoundary" in the name of that boundary. 

## What does the random string after a WebKitFormBoundary do\mean?
That's just the typical way of how a so called "boundary" between different parts of a mime structure is defined. The receiving side can tell the different parts apart by this. Same logic is used in different things, email messages too for example. Actually it is not the "random part" of that boundary that counts. The whole string is matched. It is simply a convention that each software uses a unique prefix string for such boundaries for transparency reasons. But in general the only requirement is that the chosen string must be unique throughout all contained data. Unique obviously except for the corresponding boundaries which must use exactly the same string.

