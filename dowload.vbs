dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")

' Replace the URL with your desired file location
xHttp.Open "GET", "https://github.com/Nuir4/session4.github.io/tortuga.txt", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile ".\tortuga.txt", 2 '//overwrite
end with