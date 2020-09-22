<div align="center">

## Automatic Internet Dialup


</div>

### Description

Most code snippets simply show you how to display

a connect dialog. The problem with this is that

it doesn't force a dial-up and won't alert you

when a connection is established.

This code solves those problems by using Internet

Explorer's own 'automatic dial-up' settings

(Control Panel >> Internet options >> Connections).

It utilises two little-known API calls that can

automatically connect / disconnect from the

default connection.

Note: If the 'Never Dial a Connection' option is

selected, this code will not be able to connect.

I came across this API awhile ago when my friend

suggested a forced dialup and gave me this tip.

It's actually pretty helpful.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bradley Liang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bradley-liang.md)
**Level**          |Intermediate
**User Rating**    |4.6 (73 globes from 16 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bradley-liang-automatic-internet-dialup__1-9359/archive/master.zip)

### API Declarations

```
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Declare Function InternetAutodial Lib "wininet.dll" _
(ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" _
(ByVal dwReserved As Long) As Long
```


### Source Code

```
' !! Dial the Net Automatically !!
' This waits until the connection is made and THEN
' proceeds. --Bradley Liang
Private Sub Command1_Click()
'To prompt the user to connect to the Net
If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) Then
	MsgBox "You're Connected!", vbInformation
End If
'To automatically start dialling
If InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0) Then
	MsgBox "You're Connected!", vbInformation
End If
'To disconnect an automatically dialled connection
If InternetAutodialHangup(0) Then
 MsgBox "You're Disconnected!", vbInformation
End If
End Sub
```

