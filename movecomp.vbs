'Moves computer object
'cscript movecomp.vbs <computer name> <server> <basedn> <location> <username> <password>

'Get variables
comp = WScript.Arguments.Item(0)
server = WScript.Arguments.Item(1)
basedn = WScript.Arguments.Item(2)
location = WScript.Arguments.Item(3)
username = WSCript.Arguments.Item(4)
password = WSCript.Arguments.Item(5)

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

WScript.Echo "Moving Computer "&comp&" to "&location&","&basedn

'Connect to new ou and move object
set dso = GetObject("LDAP:")
set compObj = dso.OpenDSObject("LDAP://"&server&"/"&location&","&basedn,username,password,1)

compObj.MoveHere compStr,vbNullString
WScript.Echo "Done."