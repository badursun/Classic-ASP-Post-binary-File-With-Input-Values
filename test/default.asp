<!--#include file="../fpost.asp"--><%
Dim FPost
Set FPost = New FilePost
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

    ' If you want some debug remove comment
    'Response.Write FPost.DebugBoundry()
Set FPost = Nothing
%>
