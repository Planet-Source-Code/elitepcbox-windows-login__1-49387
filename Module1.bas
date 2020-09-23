Attribute VB_Name = "Module1"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Const LOGON32_LOGON_INTERACTIVE As Long = 2
Const LOGON32_LOGON_NETWORK As Long = 3
Const LOGON32_PROVIDER_DEFAULT As Long = 0
Const LOGON32_PROVIDER_WINNT50 As Long = 3
Const LOGON32_PROVIDER_WINNT40 As Long = 2
Const LOGON32_PROVIDER_WINNT35 As Long = 1

Function VerifyLogin(sUser As String, sDomain As String, sPassword As String) As Boolean
    Dim token As Long
    VerifyLogin = LogonUser(sUser, sDomain, sPassword, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, token)
End Function

