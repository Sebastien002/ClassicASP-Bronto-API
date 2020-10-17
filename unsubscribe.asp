<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>Hello ASP</title>
</head>
<body>
<%
Dim objXMLHTTP : set objXMLHTTP = Server.CreateObject("Msxml2.XMLHTTP.3.0")
Dim strLoginRequest, strUnsubscribeRequest, strResult, strResult1, strURL, strSessionId

strURL = "http://api.bronto.com/v4"
strSessionId = "bcfe877e-b4be-4526-992d-13cf41d74cb1"

strUnsubscribeRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:hs=""http://api.bronto.com/v4"">" _
& "    <soapenv:Header>" _
& "        <hs:sessionHeader>" _
& "            <sessionId>" & strSessionId & "</sessionId>" _
& "        </hs:sessionHeader>" _
& "    </soapenv:Header>" _
& "    <soapenv:Body>" _
& "        <hs:readUnsubscribes>" _
& "            <pageNumber>1</pageNumber>" _
& "            <filter></filter>" _
& "        </hs:readUnsubscribes>" _
& "    </soapenv:Body>" _
& "</soapenv:Envelope>"

objXMLHTTP.open "post", "" & strURL & "", False

objXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
objXMLHTTP.setRequestHeader "Content-Length", Len(strUnsubscribeRequest)

'send the request and capture the result
Call objXMLHTTP.send(strUnsubscribeRequest)
strResult1 = objXMLHTTP.responseText


'display the XML
response.write strResult1
%>
</body>
</html>