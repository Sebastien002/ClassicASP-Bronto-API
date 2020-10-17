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

'URL to SOAP namespace and connection URL
strURL = "http://api.bronto.com/v4"

strLoginRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:hs=""http://api.bronto.com/v4"">" _
& "    <soapenv:Body>" _
& "        <hs:login>" _
& "            <apiToken>25498F9D-B704-493B-BEE6-DB1C79A5FFFD</apiToken>" _
& "        </hs:login>" _
& "    </soapenv:Body>" _
& "</soapenv:Envelope>"

objXMLHTTP.open "post", "" & strURL & "", False

objXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
objXMLHTTP.setRequestHeader "Content-Length", Len(strLoginRequest)

'send the request and capture the result
Call objXMLHTTP.send(strLoginRequest)
strSessionId = objXMLHTTP.responseText

response.write strSessionId

%>
</body>
</html>