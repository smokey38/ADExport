'Global variables
    Dim oContainer
    Dim OutPutFile
    Dim FileSystem
'Initialize global variables
    Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
    Set OutPutFile = FileSystem.CreateTextFile("PeopleSoftEchange.txt", True) 
    Set oContainer=GetObject("LDAP://OU=Users,OU=Warner Center,OU=Sites,OU=ES,DC=North,DC=com")  
'Enumerate Container
    EnumerateUsers oContainer
'Clean up
    OutPutFile.Close
    Set FileSystem = Nothing
    Set oContainer = Nothing
    WScript.Quit(0)
Sub EnumerateUsers(oCont)
    Dim oUser
    OutPutFile.WriteLine "Obj-Class,Display Name,Phone number,Department,Employee Number,Last Name,First Name,E-mail addresses,alias name"
    For Each oUser In oCont
        Select Case LCase(oUser.Class)
               Case "user"
                       If (Right(oUser.msExchHomeServerName,8) = "NG-ExchangeServer01") OR (Right(oUser.msExchHomeServerName,8) = "NG-ExchangeServer01") Then
                          OutPutFile.Write "Mailbox,"
                          OutPutFile.Write Chr(34) & oUser.displayName & Chr(34) & ","
                          OutPutFile.Write oUser.telephoneNumber & ","
                          OutPutFile.Write oUser.department & ","
                          OutPutFile.Write oUser.extensionAttribute2 & ","
                          OutPutFile.Write oUser.sn & ","
                          OutPutFile.Write oUser.givenName & ","
                          OutPutFile.Write "SMTP:" & oUser.mail & ","
                          OutPutFile.Write oUser.Get ("name")
                          OutPutFile.WriteLine
                       End If
               Case "organizationalunit", "container"
                    EnumerateUsers oUser
        End Select
        
    Next
End Sub

