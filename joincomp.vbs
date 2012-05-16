'Joins Computer to a domain
'cscript joincomp.vbs <new name> <server> <domain> <basedn> <ou> <username> <password>


'''
'Joining Computer to domain
'''




'Declare constants
Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const JOIN_IF_JOINED = 32

'Get parameters
strDomain = WScript.Arguments.Item(2)
strUser = WScript.Arguments.Item(5)
strPassword = WScript.Arguments.Item(6)
strOU = WScript.Arguments.Item(4)&","&WScript.Arguments.Item(3)

'Get computername 
Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")

WScript.Echo "Joining computer "&strComputer&" to domain "&strDomain

'Binding loop
Do While 1

'Attempt Join
err = objComputer.JoinDomainOrWorkGroup(strDomain, strPassword, strDomain & "\" & strUser, strOU, JOIN_DOMAIN + ACCT_CREATE + JOIN_IF_JOINED)

'Check if joined error
If err <> 0 Then
	WScript.Echo "Error: "&err
	Wscript.Quit
Else
	WScript.Echo "Computer Joined Successfully" 
End If




'''
'Searching For object
'''



'Get variables
comp = strComputer
server = WScript.Arguments.Item(1)
basedn = WScript.Arguments.Item(3)
username = strUser
password = strPassword

'Create Search Object
set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDsOObject"
ado.Properties("User ID") = username
ado.Properties("Password") = password
ado.Properties("Encrypt Password") = True
ado.Open "ldapconn"

timespent = 0
found = 0


'Searching loop
Do While timespent<30

WScript.Echo "Waiting 5 seconds..."
WScript.Sleep 5000
timespent = timespent+ 5

'Execute search
set objectList = ado.Execute("<LDAP://"&server&"/"&basedn&">;(&(ObjectClass=computer)(cn="&comp&"));ADSPath;subtree")

WScript.Echo "Checking if computer "&comp&" exists in AD"

'Make sure we found something
If objectList.RecordCount < 1 Then
	WScript.Echo "Computer Does not Exist"
Else
	WScript.Echo "Computer Exists"
	found = 1
	Exit Do
End If
Loop

If found = 1 Then
	Exit Do
End If

WScript.Echo "Computer not found in 30 Seconds, binding again"

Loop



'''
'Renaming Computer
'''


strNewName = WScript.Arguments.Item(0)

if LCase(strComputer) <> LCase(strNewName) Then

WScript.Echo "Renaming to "&strNewName

Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colComputers = objWMIService.ExecQuery _
	("Select * from Win32_ComputerSystem")

For Each objComputer in colComputers
	err = objComputer.Rename(strNewName,strPassword, strDomain & "\" & strUser)

'Check if rename error
	If err <> 0 Then
		WScript.Echo "Error: "&err
		Wscript.Quit
	Else
		WScript.Echo "Computer Renamed Successfully" 
	End If
Next

Else
	WScript.Echo "Computer already named correctly."
End If
