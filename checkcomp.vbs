'Checks if Computer object exists
'cscript checkcomp.vbs <computer name> <server> <basedn> <username> <password>

'Get variables
comp = WScript.Arguments.Item(0)
server = WScript.Arguments.Item(1)
basedn = WScript.Arguments.Item(2)
username = WSCript.Arguments.Item(3)
password = WSCript.Arguments.Item(4)

'Create Search Object
set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDsOObject"
ado.Properties("User ID") = username
ado.Properties("Password") = password
ado.Properties("Encrypt Password") = True
ado.Open "ldapconn"


'Execute search
set objectList = ado.Execute("<LDAP://"&server&"/"&basedn&">;(&(ObjectClass=computer)(cn="&comp&"));ADSPath;subtree")

WScript.Echo "Checking if computer "&comp&" exists in AD"

'Make sure we found something
If objectList.RecordCount < 1 Then
	WScript.Echo "Computer Does not Exist"
	WScript.Quit -1
Else
	WScript.Echo "Computer Exists"
	WScript.Quit 0
End If