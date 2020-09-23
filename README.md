<div align="center">

## Windows Login


</div>

### Description

Use windows current user login and password in your program, in windows 2000 and XP.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-10-23 13:12:42
**By**             |[ElitePCBOX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/elitepcbox.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Windows\_Lo16624810232003\.zip](https://github.com/Planet-Source-Code/elitepcbox-windows-login__1-49387/archive/master.zip)

### API Declarations

```
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
```





