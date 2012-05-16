'Deletes Computer object
'cscript deletecomp.vbs <computer name> <server> <basedn> <username> <password>

'Get variables
comp = WScript.Arguments.Item(0)
server = WScript.Arguments.Item(1)
basedn = WScript.Arguments.Item(2)
username = WSCript.Arguments.Item(3)
password = WSCript.Arguments.Item(4)

WScript.Echo "Deleting Computer "&comp&" from AD..."

'Create Search Object
set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDsOObject"
ado.Properties("User ID") = username
ado.Properties("Password") = password
ado.Properties("Encrypt Password") = True
ado.Open "ldapconn"


'Execute search
set objectList = ado.Execute("<LDAP://"&server&"/"&basedn&">;(&(ObjectClass=computer)(cn="&comp&"));ADSPath;subtree")


'Make sure we found something
If objectList.RecordCount < 1 Then
	WScript.Echo "Computer Does not Exist"
	WScript.Quit
Else
	compStr = objectList.Fields(0).Value
End If


'Connect to object and delete it
set dso = GetObject("LDAP:")
set compObj = dso.OpenDSObject(compStr,username,password,1)
compObj.DeleteObject(0)
WScript.Echo "Done."