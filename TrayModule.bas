Attribute VB_Name = "TrayModule"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Public SysIcon As NOTIFYICONDATA, RunningInTray As Boolean

Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long


Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
    '

    Public Const RAS95_MaxEntryName = 256
    Public Const RAS95_MaxDeviceType = 16
    Public Const RAS95_MaxDeviceName = 32
    '



Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type
    '



Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type


Public Sub ShowIcon(ByRef mail As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = mail.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = mail.Icon
    SysIcon.szTip = mail.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    RunningInTray = True
End Sub

Public Sub RemoveIcon(mail As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = mail.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = mail.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    RunningInTray = False
End Sub
Public Function IsConnected() As Boolean
    Dim TRasCon(255) As RASCONN95 'function i stole from PSC
    'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=2007
    Dim lg As Long
    Dim lpcon As Long
    Dim RetVal As Long
    Dim Tstatus As RASCONNSTATUS95
    '
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    '
    RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
    If RetVal <> 0 Then
        MsgBox "ERROR"
        Exit Function
    End If
    '
    Tstatus.dwSize = 160
    RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
    If Tstatus.RasConnState = &H2000 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

'All the encryption stuff is from a bas done by InfraRed, contact him here:
'E-Mail:  InfraRed@flashmail.com
'ICQ:  17948286 (UIN)
Function Encrypt(Start As Integer, diff As Integer, beta As Integer, alpha As Integer, times As Integer, SuperEncrypt As Boolean, text As String)
'Encrypt characters
On Error GoTo error
Dim i As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim a As Integer
SuperE = SuperEncrypt
If diff > 500 Then
diff = 500
ElseIf diff < 1 Then
diff = 1
End If
If times > 100 Then
times = 100
ElseIf times < 1 Then
times = 1
End If
If Start > 255 Then
Start = 255
ElseIf Start < 1 Then
Start = 1
End If
If beta > 5 Then
beta = 5
ElseIf beta < 1 Then
beta = 1
End If
If alpha > 5 Then
alpha = 5
ElseIf alpha < 1 Then
alpha = 1
End If
curkey = Start
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
For a = 1 To times
For i = 1 To Len(text)
    If 255 - curkey > curkey Then
    larger = 255 - curkey
    lesser = curkey
    Else
    larger = curkey
    lesser = 255 - curkey
    End If
  If Asc(Mid$(text, i, 1)) <= lesser Then
  m = Asc(Mid$(text, i, 1)) + (larger - 1)
  endstr = endstr + Chr$(m)
  Else
  m = Asc(Mid$(text, i, 1)) - lesser
  endstr = endstr + Chr$(m)
  End If
curkey = curkey + diff
  If curkey > 255 Then
  curkey = curkey - 255
  End If
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
beta = beta + (2 * diff)
alpha = alpha + diff
  If beta > 5 Then
  beta = 1
  End If
  If alpha > 5 Then
  alpha = 1
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
  If diff > 500 Then
  diff = 1
  Else
  diff = diff + diff
  End If
Next i
Text2 = ""
Text2 = endstr
endstr = ""
Next a
Encrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function Decrypt(Start As Integer, diff As Integer, beta As Integer, alpha As Integer, times As Integer, SuperEncrypt As Boolean, text As String)
'Decrypt characters
On Error GoTo error
Dim i As Integer
Dim curkey As Long
Dim m As Long
Dim endstr As String
Dim Text2 As String
Dim lesser As Double
Dim larger As Double
Dim SuperE As Boolean
Dim a As Integer
SuperE = SuperEncrypt
If diff > 500 Then
diff = 500
ElseIf diff < 1 Then
diff = 1
End If
If times > 100 Then
times = 100
ElseIf times < 1 Then
times = 1
End If
If Start > 255 Then
Start = 255
ElseIf Start < 1 Then
Start = 1
End If
If beta > 5 Then
beta = 5
ElseIf beta < 1 Then
beta = 1
End If
If alpha > 5 Then
alpha = 5
ElseIf alpha < 1 Then
alpha = 1
End If
curkey = Start
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
For a = 1 To times
For i = 1 To Len(text)
    If 255 - curkey > curkey Then
    larger = 255 - curkey
    lesser = curkey
    Else
    larger = curkey
    lesser = 255 - curkey
    End If
  If Asc(Mid$(text, i, 1)) >= larger Then
  m = Asc(Mid$(text, i, 1)) - (larger - 1)
  endstr = endstr + Chr$(m)
  Else
  m = Asc(Mid$(text, i, 1)) + lesser
  endstr = endstr + Chr$(m)
  End If
curkey = curkey + diff
  If curkey > 255 Then
  curkey = curkey - 255
  End If
curkey = (curkey * alpha) / beta
  If SuperE = True Then
    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If
  curkey = SuperEE(curkey, beta, alpha, beta)
  End If
beta = beta + (2 * diff)
alpha = alpha + diff
  If beta > 5 Then
  beta = 1
  End If
  If alpha > 5 Then
  alpha = 1
  End If
  If curkey > 255 Then
  curkey = 255 - (curkey / 255)
  ElseIf curkey < 0 Then
  curkey = 0 - (curkey / 255)
  End If
  If diff > 500 Then
  diff = 1
  Else
  diff = diff + diff
  End If
Next i
Text2 = ""
Text2 = endstr
endstr = ""
Next a
Decrypt = Text2
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Private Function SuperEE(curkey As Long, beta As Integer, alpha As Integer, times As Integer)
'For encryption:  Change the current key around more
On Error GoTo error
curkey = (((curkey / times) - (beta + times)) * alpha) + ((beta / alpha) - times)
If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
Else
curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
End If
If beta - times = 0 Then
curkey = ((curkey * alpha) + (beta * times))
Else
curkey = ((curkey * (beta - times)) + (beta - times))
  If curkey < 0 Then
  curkey = curkey + (alpha + beta)
  ElseIf curkey = 0 Then
  curkey = curkey + (alpha + times)
  Else
  curkey = curkey + (beta + times)
  End If
End If
SuperEE = curkey
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function




